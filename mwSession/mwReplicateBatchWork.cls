VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwReplicateBatchWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const RS_CL_ID = 0
Private Const RS_CL_mwrChangeTableKey = 1
Private Const RS_CL_RecordKey = 2
Private Const RS_CL_ChangeType = 3
Private Const RS_CL_mwcUsersKey = 4
Private Const RS_CL_mwcRoleTypeKey = 5
Private Const RS_CL_ChangeDateTime = 6
Private Const RS_CL_mwcSitesKey = 7
Private Const RS_CL_IsError = 8
Private Const RS_CL_mwrBatchLogOutboundKey = 9
Private Const RS_CL_OriginatingSiteKey = 10

Private Const RS_CT_ID = 11
Private Const RS_CT_Tablename = 12
Private Const RS_CT_IsActive = 13
Private Const RS_CT_mwrBatchTypeKey = 14
Private Const RS_CT_TableDescription = 15
Private Const RS_CT_SaveAuditLogs = 16
Private Const RS_CT_IsAuditTrailEnabled = 17

Private Const LONG_BIN_DEFINED_SIZE = 1073741823
Private Const BT_DA_BLOB_OUTBOUND_MODEM = 103
Private Const BT_DA_BLOB_OUTBOUND_MEDIA = 104

Private Const ZIP_IT = True
Private Const CDROM_IT = True

'
'  File Cabinet Blob Transmission
'
Private Const RS_FCB_ID = 0
Private Const RS_FCB_mwcSitesKey = 1
Private Const RS_FCB_BatchType = 2
Private Const RS_FCB_BatchKey = 3
Private Const RS_FCB_BatchNumber = 4
Private Const RS_FCB_DateMod = 5
Private Const RS_FCB_FileBlob = 6


Const BATCH_TYPE_FLEET = 1
Const BATCH_TYPE_REDIST_FLEET = 2
Const BATCH_TYPE_SITE_DATAGRAM = 100
Const BATCH_TYPE_FILECAB_MODEM = 101
Const BATCH_TYPE_FILECAB_MEDIA = 102
Const BATCH_TYPE_DA_MODEM = 103
Const BATCH_TYPE_DA_MEDIA = 104

Const DELETE_AFTER_ARCHIVE = True

Const mwrt_mwrBatchLogOutbound = 977
Const SKIP_ERRORS = True


'
' Offsets for FetchChangesColumnRS()
'
Const RS_CLC_ID = 0
Const RS_CLC_mwrChangeLogKey = 0
Const RS_CLC_ColumnName = 0
Const RS_CLC_OldValue = 0
Const RS_CLC_NewValue = 0
Const RS_CLC_OldValueLarge = 0
Const RS_CLC_NewValueLarge = 0

Const NO_ARCHIVE = True

Private Const IS_MODIFICATION = True
Private Const ARCHIVE_DATAGRAM = True
Private Const DONT_ARCHIVE = True

Private mDgGeneralDateFormat As String
Private mDgShortDateFormat As String
Private mDgDecimalSeparator As String

Private mIsNoTransactions As Boolean
Private mTotalFleetRecords As Long
Private mNewFleetBatchKey As Long
Private mIncomingOutboundBatchKey As Long

Private mStartChangeLogKey As Long
Private mEndChangeLogKey As Long

Private Const BASE_REGISTRY As String = "Software\Maritime Systems Inc\SessionSettings\"

Const SITE_TRANSACTIONS = 100

Dim mOversizeBatchLastRecKey

Dim mnBatchNo As Long
Dim mnBatchKey As Long
Dim moform As mwForm


Public Property Get IsNoTransactions() As Boolean
    IsNoTransactions = mIsNoTransactions
End Property

'Public Property Let A(ByVal Data As String)
'    mA = Data
'End Property

'Public Property Get A() As String
'    A = mA
'End Property



Public Function CreateOutboundBatchLog(BatchTypeKey As Long, ByRef BatchNo As Long, Optional SiteKey As Long) As Long
   Dim loRs As Recordset
   Dim nBatchNo As Long
   Dim ssql As String
   
   On Error GoTo FunctionError
   ' New Batch Number
   nBatchNo = GetNextBatchNumberOutbound(BatchTypeKey, SiteKey)
   If nBatchNo < 1 Then
      CreateOutboundBatchLog = -1
      Exit Function
   End If
   If SiteKey < 1 And BatchTypeKey = 100 Then
      goSession.RaiseError "Error in mwReplicateBatchWork.CreateOutboundBatchLog, no SiteKey provided for Site Transaction Type 100."
      CreateOutboundBatchLog = -1
      Exit Function
   End If
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   With loRs
      .CursorLocation = adUseClient
      ssql = "select * from mwrBatchLogOutbound where ID=-1"
      .Open ssql, goCon, adOpenDynamic, adLockOptimistic
      .AddNew
      .Fields("ID").value = goSession.MakePK("mwrBatchLogOutbound")
      CreateOutboundBatchLog = .Fields("ID").value
      .Fields("mwrBatchTypeKey").value = BatchTypeKey
      If SiteKey > 0 Then
         .Fields("mwcSitesKey").value = SiteKey
      End If
      .Fields("BatchNumber").value = nBatchNo
      If mNewFleetBatchKey > 0 Then
         .Fields("mwrBatchLogOutboundFleetKey").value = mNewFleetBatchKey
         .Fields("TotalFleetRecords").value = mTotalFleetRecords
      End If

      .Fields("StartChangeLogKey").value = mStartChangeLogKey
      .Fields("EndChangeLogKey").value = mEndChangeLogKey

      .Fields("AckBackResendRequest").value = False
      .Update
   End With
   CloseRecordset loRs
   ' pass back as ByRef value
   BatchNo = nBatchNo
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.CreateOutboundBatchLog. ", Err.Number, Err.Description
   CreateOutboundBatchLog = -1
End Function


Public Function GetNextBatchNumberOutbound(BatchTypeKey As Long, SiteKey As Long) As Long
   Dim ssql As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   If SiteKey > 0 Then
      ssql = "select max(BatchNumber) from mwrBatchLogOutbound where mwcSitesKey=" & SiteKey & _
        " and mwrBatchTypeKey=" & BatchTypeKey
   Else
      ssql = "select max(BatchNumber) from mwrBatchLogOutbound where mwrBatchTypeKey=" & BatchTypeKey
   End If
   Set loRs = New Recordset
   With loRs
      .CursorLocation = adUseClient
      .Open ssql, goCon, adOpenDynamic, adLockOptimistic
      If IsNull(.Fields(0).value) Then
         GetNextBatchNumberOutbound = 1
      Else
         GetNextBatchNumberOutbound = .Fields(0).value + 1
      End If
   End With
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.GetNextBatchNumberOutbound. ", Err.Number, Err.Description
   GetNextBatchNumberOutbound = -1
End Function

Private Function FetchDetailRS(TableName As String, RecordKey As Long) As Recordset
   Dim ssql As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   ssql = "select * from " & TableName & " where ID=" & RecordKey
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open ssql, goCon, adOpenForwardOnly, adLockReadOnly
   Set FetchDetailRS = loRs
   Set loRs = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.FetchDetailRS. ", Err.Number, Err.Description
End Function

Private Function FetchChangesRS(BatchTypeKey As Long, FleetBatchKey As Long, Optional SiteKey As Long, Optional MinChangeLogKey As Long, Optional MaxChangeLogKey As Long) As Recordset
   Dim ssql As String
   Dim sAndQry As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   
   If goSession.ThisSite.IsLogIncomingTransactions Then
      sAndQry = " AND (IsLoggedTransaction Is Null or IsLoggedTransaction = 0) "
   End If
'   If SiteKey > 0 Then
'      sSQL = "SELECT mwrChangeLog.*, mwrChangeTable.* " & _
'        "FROM mwrChangeTable INNER JOIN mwrChangeLog ON mwrChangeTable.ID = mwrChangeLog.mwrChangeTableKey " & _
'        "where mwcSitesKey=" & SiteKey & " and mwrBatchTypeKey=" & BatchTypeKey & _
'        " and mwrBatchLogOutboundKey is null order by mwrChangeLog.ID"
'   Else
'      sSQL = "SELECT mwrChangeLog.*, mwrChangeTable.* " & _
'        "FROM mwrChangeTable INNER JOIN mwrChangeLog ON mwrChangeTable.ID = mwrChangeLog.mwrChangeTableKey " & _
'        "where mwrBatchTypeKey=" & BatchTypeKey & _
'        " and mwrBatchLogOutboundKey is null order by mwrChangeLog.ID"
'   End If
   ssql = "SELECT mwrChangeLog.ID,mwrChangeLog.mwrChangeTableKey,mwrChangeLog.RecordKey,mwrChangeLog.ChangeType,mwrChangeLog.mwcUsersKey, " & _
      " mwrChangeLog.mwcRoleTypeKey,mwrChangeLog.ChangeDateTime,mwrChangeLog.mwcSitesKey,mwrChangeLog.IsError, " & _
      " mwrChangeLog.mwrBatchLogOutboundKey,mwrChangeLog.OriginatingSiteKey,mwrChangetable.ID as ChangeTableID, " & _
      " mwrChangetable.Tablename,mwrChangetable.IsActive,mwrChangetable.mwrBatchTypeKey,mwrChangetable.TableDescription, " & _
      " mwrChangetable.SaveAuditLogs,mwrChangetable.IsAuditTrailEnabled " & _
      " FROM mwrChangeTable , mwrChangeLog " & _
      " WHERE mwrChangeTable.ID = mwrChangeLog.mwrChangeTableKey AND (("

   If SiteKey > 0 Then
      ssql = ssql & " mwcSitesKey = " & SiteKey & " AND "
   End If
   
   If BatchTypeKey = 1 Then
      ssql = ssql & " mwrBatchTypeKey IN(1, 2) AND mwrBatchLogOutboundKey Is Null " & sAndQry & " ) "
   Else
      ssql = ssql & " mwrBatchTypeKey = " & BatchTypeKey & " AND " & _
            " mwrBatchLogOutboundKey Is Null " & sAndQry & " ) "
   End If
   
   If FleetBatchKey > 0 Then
      ssql = ssql & " OR (mwrChangeLog.mwrBatchLogOutboundKey = " & FleetBatchKey
      
      ' Are we on the SHORE sending records to another site? If so then don't send records that originated on that site
      ' because the site already has those records.
      
      ' If we are on a SHIP then the SiteKey passed in is that ship's site and we DO need to send those records.
      
      If goSession.Site.SiteType = SITE_TYPE_SHORE And SiteKey > 0 Then
         ssql = ssql & " AND (mwrChangeLog.OriginatingSiteKey Is Null OR mwrChangeLog.OriginatingSiteKey <> " & SiteKey & ")"
      End If
      ssql = ssql & ")"
   End If
   
   If MinChangeLogKey > 0 Then
      ssql = ssql & ") AND (mwrChangeLog.ID >= " & MinChangeLogKey
   End If
   
   If MaxChangeLogKey > 0 Then
      ssql = ssql & ") AND (mwrChangeLog.ID <= " & MaxChangeLogKey
   End If
   
   ssql = ssql & ") ORDER BY mwrChangeLog.ID"

   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open ssql, goCon, adOpenDynamic, adLockOptimistic
   loRs.Properties("Unique Table").value = "mwrChangeLog"
   If Not (loRs.EOF Or loRs.BOF) Then
      loRs.MoveFirst
   End If
   Set FetchChangesRS = loRs
   Set loRs = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.FetchChangesRS. ", Err.Number, Err.Description
End Function

Private Function FetchAuditRS(BatchKey As Long, FleetBatchKey As Long, Optional MinChangeLogKey As Long, _
                              Optional MaxChangeLogKey As Long) As Recordset
   Dim ssql As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   
   ssql = "SELECT mwrAuditLog.ID,mwrAuditLog.mwrChangeTableKey,mwrAuditLog.RecordKey,mwrAuditLog.ChangeType, " & _
      " mwrAuditLog.mwcUsersKey,mwrAuditLog.mwcRoleTypeKey,mwrAuditLog.ChangeDateTime,mwrAuditLog.mwcSitesKey, " & _
      " mwrAuditLog.IsError,mwrAuditLog.mwrBatchLogOutboundKey,mwrAuditLog.OriginatingSiteKey, " & _
      " mwrChangetable.ID as ChangeTableID,mwrChangetable.Tablename,mwrChangetable.IsActive,mwrChangetable.mwrBatchTypeKey," & _
      " mwrChangetable.TableDescription,mwrChangetable.SaveAuditLogs,mwrChangetable.IsAuditTrailEnabled " & _
      " FROM mwrChangeTable , mwrAuditLog " & _
      " WHERE mwrChangeTable.ID = mwrAuditLog.mwrChangeTableKey AND (mwrBatchLogOutboundKey=" & _
      BatchKey & " or mwrBatchLogOutboundKey=" & FleetBatchKey & ") "
      
   If MinChangeLogKey > 0 Then
      ssql = ssql & " AND mwrAuditLog.ID >= " & MinChangeLogKey
   End If
   
   If MaxChangeLogKey > 0 Then
      ssql = ssql & " AND mwrAuditLog.ID <= " & MaxChangeLogKey
   End If
   

   ssql = ssql & " ORDER BY mwrAuditLog.ID"
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open ssql, goCon, adOpenDynamic, adLockOptimistic
   loRs.Properties("Unique Table").value = "mwrAuditLog"
   If Not (loRs.EOF Or loRs.BOF) Then
      loRs.MoveFirst
   End If
   Set FetchAuditRS = loRs
   Set loRs = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.FetchAuditRS. ", Err.Number, Err.Description
End Function



Private Function FetchSatelliteChangesRS() As Recordset
   Dim ssql As String
   Dim loRs As Recordset
   Dim loKeys As ConfigKeys
   Dim sAndQry As String
   On Error GoTo FunctionError
   
   If goSession.ThisSite.IsLogIncomingTransactions Then
      sAndQry = " AND (IsLoggedTransaction Is Null or IsLoggedTransaction = 0) "
   End If

   Set loKeys = goSession.GetEventSecurityKeys(MW_EVENT_WORKFLOW_AGENT)
   If loKeys.GetBoolKeyValue("SendBatchTypeNotNull") Then
      ssql = "SELECT mwrChangeLog.ID,mwrChangeLog.mwrChangeTableKey,mwrChangeLog.RecordKey,mwrChangeLog.ChangeType, " & _
         " mwrChangeLog.mwcUsersKey,mwrChangeLog.mwcRoleTypeKey,mwrChangeLog.ChangeDateTime,mwrChangeLog.mwcSitesKey, " & _
         " mwrChangeLog.IsError,mwrChangeLog.mwrBatchLogOutboundKey,mwrChangeLog.OriginatingSiteKey, " & _
         " mwrChangetable.ID as ChangeTableID,mwrChangetable.Tablename,mwrChangetable.IsActive,mwrChangetable.mwrBatchTypeKey," & _
         " mwrChangetable.TableDescription,mwrChangetable.SaveAuditLogs,mwrChangetable.IsAuditTrailEnabled " & _
         " FROM mwrChangeTable , mwrChangeLog " & _
         " WHERE mwrChangeTable.ID = mwrChangeLog.mwrChangeTableKey " & _
         " AND mwrBatchLogOutboundKey Is Null " & sAndQry & " and mwrChangeTable.mwrBatchTypeKey is not null" & _
         " ORDER BY mwrChangeLog.ID"
   Else
      ssql = "SELECT mwrChangeLog.ID,mwrChangeLog.mwrChangeTableKey,mwrChangeLog.RecordKey,mwrChangeLog.ChangeType, " & _
         " mwrChangeLog.mwcUsersKey,mwrChangeLog.mwcRoleTypeKey,mwrChangeLog.ChangeDateTime,mwrChangeLog.mwcSitesKey, " & _
         " mwrChangeLog.IsError,mwrChangeLog.mwrBatchLogOutboundKey,mwrChangeLog.OriginatingSiteKey, " & _
         " mwrChangetable.ID as ChangeTableID,mwrChangetable.Tablename,mwrChangetable.IsActive,mwrChangetable.mwrBatchTypeKey," & _
         " mwrChangetable.TableDescription,mwrChangetable.SaveAuditLogs,mwrChangetable.IsAuditTrailEnabled " & _
         " FROM mwrChangeTable , mwrChangeLog " & _
         " WHERE mwrChangeTable.ID = mwrChangeLog.mwrChangeTableKey " & _
         " AND mwrBatchLogOutboundKey Is Null " & sAndQry & " " & _
         " ORDER BY mwrChangeLog.ID"
   End If

   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open ssql, goCon, adOpenDynamic, adLockOptimistic
   loRs.Properties("Unique Table").value = "mwrChangeLog"
   If Not (loRs.EOF Or loRs.BOF) Then
      loRs.MoveFirst
   End If
   Set FetchSatelliteChangesRS = loRs
   Set loRs = Nothing
   Set loKeys = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.FetchSatelliteChangesRS. ", Err.Number, Err.Description
End Function

Public Function FetchReplicationSites(Optional SiteKey As Long) As Recordset
   Dim ssql As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   ssql = "select * from mwcSites where IsReplicateSite Is Not Null and IsReplicateSite<>0"
   If SiteKey > 0 Then
      ' ship - include oneself
      ssql = ssql & " and ID=" & SiteKey
   Else
      ' Shore - exclude thyself
      ssql = ssql & " and ID<>" & goSession.Site.SiteKey
   End If
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open ssql, goCon, adOpenDynamic, adLockOptimistic
   Set FetchReplicationSites = loRs
   Set loRs = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.FetchReplicationSites. ", Err.Number, Err.Description
End Function


Private Function FetchDetailModifyRS(TableName As String, RecordKey As Long, ChangeLogKey As Long) As Recordset
   Dim ssql As String
   Dim loRsColumns As Recordset
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRsColumns = New Recordset
   loRsColumns.CursorLocation = adUseClient
   ssql = "select * from mwChangeLogColumn where mwrChangeLogKey=" & ChangeLogKey & " ORDER BY ID"
   loRsColumns.Open ssql, goCon, adOpenForwardOnly, adLockReadOnly
   If loRsColumns.RecordCount < 1 Then
      goSession.RaiseWorkflowStatus "Warning in mwReplicateWork.FetchDetailModifyRS, Table Modify with no columns: " & ChangeLogKey
      CloseRecordset loRsColumns
      Exit Function
   End If
   ssql = "select ID"
   Do While Not loRsColumns.EOF
      ssql = ssql & "," & loRsColumns!ColumnName
      loRsColumns.MoveNext
   Loop
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open ssql, goCon, adOpenForwardOnly, adLockReadOnly
   Set FetchDetailModifyRS = loRs
   Set loRs = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.FetchDetailModifyRS. ", Err.Number, Err.Description
End Function

Public Function FinalizeOutboundBatchlog(BatchKey As Long, DatagramFormID As String, TotalRecords As Long, Optional mNewFleetBatchKey As Long) As Boolean
   Dim loRs As Recordset
   Dim ssql As String
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   ssql = "select * from mwrBatchLogOutbound where ID=" & BatchKey
   loRs.Open ssql, goCon, adOpenDynamic, adLockOptimistic
   If loRs.RecordCount < 1 Then
      goSession.RaiseError "Error in mwReplicateBatchWork.FinalizeOutboundBatchlog, Batch record missing: " & _
       BatchKey & "while datagram, " & DatagramFormID & ", has been sent !"
      FinalizeOutboundBatchlog = False
      CloseRecordset loRs
      Exit Function
   End If
   With loRs
      If Len(DatagramFormID) > 0 Then
         .Fields("DatagramFormID").value = DatagramFormID
      End If
      .Fields("TotalRecords").value = TotalRecords
      
      If .Fields("mwrBatchTypeKey").value = BATCH_TYPE_FILECAB_MODEM Or .Fields("mwrBatchTypeKey").value = BATCH_TYPE_FILECAB_MEDIA _
       Or .Fields("mwrBatchTypeKey").value = BATCH_TYPE_DA_MODEM Or .Fields("mwrBatchTypeKey").value = BATCH_TYPE_DA_MEDIA Then
         .Fields("DateSent").value = Now()
      End If
      
      .Fields("DatePrepared").value = Now()
      
      If mNewFleetBatchKey > 0 Then
         .Fields("mwrBatchLogOutboundFleetKey").value = mNewFleetBatchKey
      End If
      
      If mStartChangeLogKey > 0 Then
         .Fields("StartChangeLogKey").value = mStartChangeLogKey
      End If
      
      If mEndChangeLogKey > 0 Then
         .Fields("EndChangeLogKey").value = mEndChangeLogKey
      End If
      
      .Update
   End With
   CloseRecordset loRs
   FinalizeOutboundBatchlog = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.FinalizeOutboundBatchlog. ", Err.Number, Err.Description
   FinalizeOutboundBatchlog = False
End Function
Public Function FetchOutboundBatchlogRS(BatchType As Long) As Recordset
   Dim loRs As Recordset
   Dim ssql As String
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   ssql = "select * from mwrBatchLogOutbound WHERE DateSent IS NULL and mwrBatchTypeKey=" & BatchType & _
          " AND mwcSitesKey IN (SELECT ID FROM mwcSites WHERE " & _
          " (IsSuspendTransmit=0 or IsSuspendTransmit is null) " & _
          " AND (IsReplicateSite<>0 and IsReplicateSite is not null)) "
   
   ' include resend mwrBatchLogOutBound records also
   
'   sSQL = "SELECT mwrBatchLogOutbound.*" & _
'    " FROM mwcSites INNER JOIN mwrBatchLogOutbound ON mwcSites.ID = mwrBatchLogOutbound.mwcSitesKey " & _
'    " WHERE mwrBatchLogOutbound.DateSent Is Null AND mwrBatchLogOutbound.mwrBatchTypeKey=100 " & _
'    " AND (mwcSites.IsSuspendTransmit=0 or mwcSites.IsSuspendTransmit is null) " & _
'    " AND (mwcSites.IsReplicateSite<>0 and  mwcSites.IsReplicateSite is not null) "

  
   loRs.Open ssql, goCon, adOpenDynamic, adLockOptimistic
   
   Set FetchOutboundBatchlogRS = loRs
   Set loRs = Nothing
   
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.FetchOutboundBatchlogRS. ", Err.Number, Err.Description
End Function

Public Function CreateBatchDatagrams(BatchTypeKey As Long, Optional SiteKey As Long, Optional IsDeferShoreAckUntilMoreDataPending As Boolean, Optional IsNoPendingTransactionFound As Boolean) As Boolean
   Dim ssql As String
   Dim loRsSites As Recordset
   Dim loRsChanges As Recordset
   Dim loRs As Recordset
   Dim loRsOBW As Recordset
   Dim loRec As ConfigKeys
'   Dim moform As mwForm
   Dim loDbWork As mwDataWork
'   Dim strTemplate As String
   Dim sTargetSiteID As String
   Dim IsError As Boolean
   Dim sEmailAddress As String
'   Dim sFormID As String
   Dim s As String
   Dim sArcFile As String
   Dim fso As FileSystemObject
   Dim RecordsWritten As Long
   Dim MaxRecordsForSite As Long
   Dim IsArchiving As Boolean
   Dim strArchivingCmd As String
   Dim IsContinue As Boolean
   Dim IsAuditTrailEnabledForSite As Boolean
   Dim loAuditTrail As snAuditTrail

   On Error GoTo FunctionError
   '
   '
   '
   If SiteKey > 0 Then
      MsgBox "SiteKey not supported"
   End If
   '
   ' Fetch Regional Settings
   '
   goSession.RaiseWorkflowStatus "Send Database Changes: Creating outbound replication datagrams."
   goSession.MemoryLogIt "CreateBatchDatagrams "
   
   FixChangeLogSiteKeys
   
   mDgGeneralDateFormat = DATAGRAM_DATE_FORMAT
   mDgShortDateFormat = DATAGRAM_SHORT_DATE_FORMAT
   mDgDecimalSeparator = DATAGRAM_DECIMAL_FORMAT
   
   mOversizeBatchLastRecKey = 0
   
   Set loRsOBW = FetchOutboundWork
   
   If loRsOBW Is Nothing Then
      goSession.RaiseWorkflowStatus "Could not open mwrOutboundWork table. Operation Cancelled."
      CreateBatchDatagrams = False
      Exit Function
   End If
   
   If loRsOBW.RecordCount < 1 Then
   
      mStartChangeLogKey = GetMinChangeLogKey
      mEndChangeLogKey = GetMaxChangeLogKey
      
      
      If mEndChangeLogKey < 1 Then
         goSession.RaiseWorkflowStatus "No transactions found, Operation Cancelled."
         CloseRecordset loRsOBW
         IsNoPendingTransactionFound = True
         CreateBatchDatagrams = False
         Exit Function
      End If
   
      '
      ' Step 1 Fetch Sites to replicate to...
      '
      If goSession.Site.SiteType = SITE_TYPE_SHIP Then
         goSession.RaiseWorkflowStatus "Starting Create Batch Datagrams process for this site going ashore..."
         Set loRsSites = FetchReplicationSites(goSession.Site.SiteKey)
      Else
         goSession.RaiseWorkflowStatus "Starting Create Batch Datagrams process for all sites. "
         Set loRsSites = FetchReplicationSites()
      End If
      '
      '
      '
      ' For all replicated sites...
      '
      goSession.MemoryLogIt "CreateBatchDatagrams 2 "
      
      goCon.BeginTrans
         goSession.RaiseWorkflowStatus "...Begin Transaction: Making mwrOutboundwork records."
      
         goSession.RaiseWorkflowStatus "......MarkFleetBatchDatagrams"
     
         '
         '...make the fleet batch
         If goSession.Site.SiteType = SITE_TYPE_SHORE Then
            mNewFleetBatchKey = MarkFleetBatchDatagrams(1)
         Else
            mNewFleetBatchKey = MarkFleetBatchDatagrams(2)
         End If
         
         If mNewFleetBatchKey < 0 Then
            CloseRecordset loRsOBW
            CloseRecordset loRsSites
            goCon.RollbackTrans
            CreateBatchDatagrams = False
            Exit Function
         End If
         
         goSession.RaiseWorkflowStatus "......Populating mwrOutboundwork records."
         
         Do While Not loRsSites.EOF
         
            loRsOBW.AddNew
            loRsOBW!FleetBatchKey = mNewFleetBatchKey
            loRsOBW!mwcSitesKey = loRsSites!ID
            loRsOBW!mwrBatchTypeKey = BatchTypeKey
            loRsOBW!StartChangeLogKey = mStartChangeLogKey
            loRsOBW!EndChangeLogKey = mEndChangeLogKey
            loRsOBW!IsSiteDone = False
            loRsOBW.Update
            
            loRsSites.MoveNext
         Loop
         
         CloseRecordset loRsOBW
         CloseRecordset loRsSites
         
      goCon.CommitTrans
      goSession.RaiseWorkflowStatus "...Commit Transaction: mwrOutboundwork populated."
   
      Set loRsOBW = FetchOutboundWork
      
      If loRsOBW Is Nothing Then
         goSession.RaiseWorkflowStatus "Could not open mwrOutboundWork table. Operation Cancelled."
         CreateBatchDatagrams = False
         Exit Function
      End If
   Else
      goSession.RaiseWorkflowStatus "Found " & loRsOBW.RecordCount & " mwrOutboundWork records. Restarting Datagram creation with site " & goSession.Site.GetSiteName(loRsOBW!mwcSitesKey)
   End If
   
   goSession.MemoryLogIt "CreateBatchDatagrams 3 "
   
   goSession.RaiseWorkflowStatus "...Beginning main loop"
'   goSession.RaiseWorkflowStatus "......FleetBatchKey = " & mNewFleetBatchKey
'   goSession.RaiseWorkflowStatus "......StartChangeLogKey = " & mStartChangeLogKey
'   goSession.RaiseWorkflowStatus "......EndChangeLogKey = " & mEndChangeLogKey
   
   Do While loRsOBW.RecordCount > 0
      
      CloseRecordset loRsChanges
      CloseRecordset loRsSites
      
'      If IsUseReadUncommitted Then
'         goCon.IsolationLevel = adXactReadUncommitted
'         goSession.RaiseWorkflowStatus "...Using Isolation Level Read Uncommitted  "
'
'      Else
'         goCon.IsolationLevel = adXactReadCommitted
'         goSession.RaiseWorkflowStatus "...Using Isolation Level Read Committed  "
'      End If
      
      goCon.BeginTrans
         goSession.RaiseWorkflowStatus "..."
         goSession.RaiseWorkflowStatus "..."
         goSession.RaiseWorkflowStatus "...Begin Transaction: "
         goSession.MemoryLogIt "CreateBatchDatagrams 4 "
   
         mNewFleetBatchKey = loRsOBW!FleetBatchKey
         mStartChangeLogKey = loRsOBW!StartChangeLogKey
         mEndChangeLogKey = loRsOBW!EndChangeLogKey

         If goSession.Site.SiteType = SITE_TYPE_SHIP Then
            IsAuditTrailEnabledForSite = True
         Else
            Set loAuditTrail = New snAuditTrail
            IsAuditTrailEnabledForSite = loAuditTrail.GetAuditEnabledStatusForSite(loRsOBW!mwcSitesKey)
            KillObject loAuditTrail
         End If
         
         Set loRsSites = FetchReplicationSites(loRsOBW!mwcSitesKey)
         
         MaxRecordsForSite = ZeroNull(loRsSites!MaxRecordsPerDatagram)
         
         sArcFile = ""
         goSession.RaiseWorkflowStatus "...Creating Batch Datagram for: " & loRsSites!SiteName & ", MaxRecordsPerDatagram = " & MaxRecordsForSite
'         goSession.RaiseWorkflowStatus "......FleetBatchKey = " & mNewFleetBatchKey
'         goSession.RaiseWorkflowStatus "......StartChangeLogKey = " & mStartChangeLogKey
'         goSession.RaiseWorkflowStatus "......EndChangeLogKey = " & mEndChangeLogKey
         DoEvents
         '
         ' Fetch changes for the site...
         '
         Set loRsChanges = FetchChangesRS(BatchTypeKey, mNewFleetBatchKey, loRsSites!ID, mStartChangeLogKey, mEndChangeLogKey)
         
         DoEvents
         If loRsChanges Is Nothing Then
            goSession.RaiseError "Error in mwReplicateBatchWork.CreateBatchDatagrams, unable to open changes table. Operation Cancelled."
            CloseRecordset loRsChanges
            CloseRecordset loRsSites
            CloseRecordset loRsOBW
            CreateBatchDatagrams = False
            goCon.RollbackTrans
            Exit Function
         End If
         '
         '
         ' Any Changes  for this ship
         '
         IsContinue = False
         If loRsChanges.RecordCount > 0 Then
            IsContinue = True
         End If
         
         If BatchTypeKey = SITE_TRANSACTIONS And IsDeferShoreAckUntilMoreDataPending = True Then
            IsContinue = False
            Do While Not loRsChanges.EOF
               If UCase(loRsChanges!TableName) <> "MWRBATCHLOGOUTBOUND" Then
                  IsContinue = True
                  Exit Do
               End If
               loRsChanges.MoveNext
            Loop
            If Not loRsChanges.BOF Then
               loRsChanges.MoveFirst
            End If
         End If
         
         goSession.MemoryLogIt "CreateBatchDatagrams 5 "
         
         If IsNull(loRsSites!ShortSiteID) Then
            goSession.RaiseError "Error in mwReplicateBatchWork.CreateBatchDatagrams: Site " & loRsSites!SiteName & " has a Null ShortSiteID, can not generate datagrams for that site."
         ElseIf IsNull(loRsSites!SiteID) Then
            goSession.RaiseError "Error in mwReplicateBatchWork.CreateBatchDatagrams: Site " & loRsSites!SiteName & " has a Null SiteID, can not generate datagrams for that site."
'         ElseIf loRsChanges.RecordCount > 0 Then
         ElseIf IsContinue = True Then
            'goCon.BeginTrans
            '
            ' OK, Stage the Form...
            '
            
            mStartChangeLogKey = 0
            mEndChangeLogKey = 0
            
            If PrepareNewDatagram(BatchTypeKey, loRsSites!ID, loRsSites!ShortSiteID) = False Then
               CloseRecordset loRsChanges
               CloseRecordset loRsSites
               CreateBatchDatagrams = False
               goCon.RollbackTrans
               Exit Function
            End If
            
            
            '
            ' Populate Target Info in Header...
            '
            
            If WriteHeaderRecord(BatchTypeKey, loRsSites!SiteID, loRsSites!ID) = False Then
               CloseRecordset loRsChanges
               CloseRecordset loRsSites
               CreateBatchDatagrams = False
               goCon.RollbackTrans
               Exit Function
            End If
            
            '
            ' Loop through change records...
            '
            If goSession.Site.SiteType = SITE_TYPE_SHORE Then
               goSession.RaiseWorkflowStatus "...Writing " & loRsChanges.RecordCount & " Changes to " & loRsSites!SiteName
            Else
               goSession.RaiseWorkflowStatus "...Writing " & loRsChanges.RecordCount & " Changes to " & goSession.Site.TargetReplicateSiteID
            End If
            
            RecordsWritten = 0
            goSession.MemoryLogIt "CreateBatchDatagrams 6 "
            
            Do While Not loRsChanges.EOF
               '
               DoEvents
               
               If MaxRecordsForSite > 0 And RecordsWritten >= MaxRecordsForSite Then
                  goSession.RaiseWorkflowStatus "...Max Records Per Datagram reached, starting new datagram for " & loRsSites!SiteName

                  If CloseOldDatagram(BatchTypeKey, loRsSites!SiteID, loRsSites!ID, loRsSites!ShortSiteID) = False Then
                     CloseRecordset loRsChanges
                     CloseRecordset loRsSites
                     If Not moform Is Nothing Then
                        moform.TearDown
                     End If
                     CreateBatchDatagrams = False
                     goCon.RollbackTrans
                     Exit Function
                  End If

                  goSession.RaiseWorkflowStatus "...Previous datagram closed and copied to ArchivedOK, new datagram started"
'                  goSession.RaiseWorkflowStatus "......New StartChangeLogKey = " & loRsChanges.Fields(RS_CL_ID).value
                  goSession.RaiseWorkflowStatus "..."

                  loRsOBW!StartChangeLogKey = loRsChanges.Fields(RS_CL_ID).value
                  loRsOBW.Update
                  goCon.CommitTrans
                  goCon.BeginTrans

                  If StartNewDatagram(BatchTypeKey, loRsSites!SiteID, loRsSites!ID, loRsSites!ShortSiteID) = False Then
                     CloseRecordset loRsChanges
                     CloseRecordset loRsSites
                     If Not moform Is Nothing Then
                        moform.TearDown
                     End If
                     CreateBatchDatagrams = False
                     goCon.RollbackTrans
                     Exit Function
                  End If

                  RecordsWritten = 0

                  DoEvents
               End If
               
               Select Case loRsChanges.Fields(RS_CL_ChangeType).value
                  
                  Case "L"
                  
                     ' Replicated Log (like Schema) request site to send user logfile
                     
                     goSession.RaiseWorkflowStatus "...Generating Remote Log Datagram"
   
                     If RecordsWritten > 0 Then
                        If CloseOldDatagram(BatchTypeKey, loRsSites!SiteID, loRsSites!ID, loRsSites!ShortSiteID) = False Then
                           CloseRecordset loRsChanges
                           CloseRecordset loRsSites
                           If Not moform Is Nothing Then
                              moform.TearDown
                           End If
                           CreateBatchDatagrams = False
                           goCon.RollbackTrans
                           Exit Function
                        End If
                        
                        goSession.RaiseWorkflowStatus "..."
                        
                        If StartNewDatagram(BatchTypeKey, loRsSites!SiteID, loRsSites!ID, loRsSites!ShortSiteID) = False Then
                           CloseRecordset loRsChanges
                           CloseRecordset loRsSites
                           If Not moform Is Nothing Then
                              moform.TearDown
                           End If
                           CreateBatchDatagrams = False
                           goCon.RollbackTrans
                           Exit Function
                        End If
                        
                     End If
                     
                     RecordsWritten = 0
                     
                     ' log file is established as site datagram
                     Set loRec = moform.DataForm.AddRecord
                     loRec.Add "OPERATION", "L", "OPERATION"
                     loRec.Add "TABLE", loRsChanges.Fields(RS_CT_Tablename).value, "TABLE"
                     'loRec.Add "mwcSiteKey", loRsChanges.Fields(RS_CL_mwcSitesKey).value, "MWCSITESKEY"
                     loRec.Add "UserKey", loRsChanges.Fields(RS_CL_mwcUsersKey).value, "USERKey"
                     If Not WriteChangeColumns(loRec, loRsChanges.Fields(RS_CL_ID).value) Then
                        CloseRecordset loRsSites
                        CloseRecordset loRsChanges
                        goSession.RaiseWorkflowStatus "Error Creating Batch Datagrams. Process cancelled (Log)."
                        goCon.RollbackTrans
                        Exit Function
                     End If
                        
                  Case "S"
                     ' Replicated Schema
                     goSession.RaiseWorkflowStatus "...Generating Remote Schema Datagram"
   
                     If RecordsWritten > 0 Then
                        If CloseOldDatagram(BatchTypeKey, loRsSites!SiteID, loRsSites!ID, loRsSites!ShortSiteID) = False Then
                           CloseRecordset loRsChanges
                           CloseRecordset loRsSites
                           If Not moform Is Nothing Then
                              moform.TearDown
                           End If
                           CreateBatchDatagrams = False
                           goCon.RollbackTrans
                           Exit Function
                        End If
                        
'                        goSession.RaiseWorkflowStatus "......New StartChangeLogKey = " & loRsChanges.Fields(RS_CL_ID).value
                        goSession.RaiseWorkflowStatus "..."
                        
                        loRsOBW!StartChangeLogKey = loRsChanges.Fields(RS_CL_ID).value
                        loRsOBW.Update
                        goCon.CommitTrans
                        goCon.BeginTrans
                        
                        If StartNewDatagram(BatchTypeKey, loRsSites!SiteID, loRsSites!ID, loRsSites!ShortSiteID) = False Then
                           CloseRecordset loRsChanges
                           CloseRecordset loRsSites
                           If Not moform Is Nothing Then
                              moform.TearDown
                           End If
                           CreateBatchDatagrams = False
                           goCon.RollbackTrans
                           Exit Function
                        End If
                        
                     End If
                     
                     RecordsWritten = 0
                     
                     Set loRec = moform.DataForm.AddRecord
                     loRec.Add "OPERATION", "S", "OPERATION"
                     loRec.Add "TABLE", loRsChanges.Fields(RS_CT_Tablename).value, "TABLE"
                     If Not WriteChangeColumns(loRec, loRsChanges.Fields(RS_CL_ID).value) Then
                        CloseRecordset loRsSites
                        CloseRecordset loRsChanges
                        goSession.RaiseWorkflowStatus "Error Creating Batch Datagrams. Process cancelled."
                        goCon.RollbackTrans
                        Exit Function
                     End If
                     
                  Case "D"
                     ' Delete Transaction
                     With moform.DataForm.AddRecord
                        .Add "OPERATION", "D", "OPERATION"
                        If BoolNull(loRsChanges.Fields(RS_CT_IsAuditTrailEnabled).value) = True And IsAuditTrailEnabledForSite Then
                           .Add "CHANGEDATE", MakeDatagramDate(loRsChanges.Fields(RS_CL_ChangeDateTime).value), "CHANGEDATE"
                           .Add "USERKEY", loRsChanges.Fields(RS_CL_mwcUsersKey).value, "USER"
                        End If

                        .Add "TABLE", loRsChanges.Fields(RS_CT_Tablename).value, "TABLE"
                        .Add "ID", loRsChanges.Fields(RS_CL_RecordKey).value, "ID"
                     End With
                  Case "A"
                     ' Add Transaction
                     Set loRec = moform.DataForm.AddRecord
                     loRec.Add "OPERATION", "A", "OPERATION"
                     If BoolNull(loRsChanges.Fields(RS_CT_IsAuditTrailEnabled).value) = True And IsAuditTrailEnabledForSite Then
                        loRec.Add "CHANGEDATE", MakeDatagramDate(loRsChanges.Fields(RS_CL_ChangeDateTime).value), "CHANGEDATE"
                        loRec.Add "USERKEY", loRsChanges.Fields(RS_CL_mwcUsersKey).value, "USER"
                     End If

                     loRec.Add "TABLE", loRsChanges.Fields(RS_CT_Tablename).value, "TABLE"
                     If Not WriteChangeColumns(loRec, loRsChanges.Fields(RS_CL_ID).value) Then
                        CloseRecordset loRsSites
                        CloseRecordset loRsChanges
                        goSession.RaiseWorkflowStatus "Error Creating Batch Datagrams. Process cancelled."
                        goCon.RollbackTrans
                        Exit Function
                     End If
                  Case "M"
                     ' Modify Transaction
                     Set loRec = moform.DataForm.AddRecord
                     loRec.Add "OPERATION", "M", "OPERATION"
                     If BoolNull(loRsChanges.Fields(RS_CT_IsAuditTrailEnabled).value) = True And IsAuditTrailEnabledForSite Then
                        loRec.Add "CHANGEDATE", MakeDatagramDate(loRsChanges.Fields(RS_CL_ChangeDateTime).value), "CHANGEDATE"
                        loRec.Add "USERKEY", loRsChanges.Fields(RS_CL_mwcUsersKey).value, "USER"
                     End If

                     loRec.Add "TABLE", loRsChanges.Fields(RS_CT_Tablename).value, "TABLE"
                     loRec.Add "ID", loRsChanges.Fields(RS_CL_RecordKey).value, "ID"
                     If Not WriteChangeColumns(loRec, loRsChanges.Fields(RS_CL_ID).value) Then
                        CloseRecordset loRsSites
                        CloseRecordset loRsChanges
                        goSession.RaiseWorkflowStatus "Error Creating Batch Datagrams. Process cancelled."
                        goCon.RollbackTrans
                        Exit Function
                     End If
                  Case Else
                     goSession.RaiseWorkflowStatus "Warning in mwReplicateBatchWork. invalid change type code. Key: " & loRsChanges.Fields(RS_CL_ID).value
               End Select
               '
               ' Update BatchKey in Change Log record if it isn't already set to the Fleet Batch Key
               '
               If IsNull(loRsChanges.Fields(RS_CL_mwrBatchLogOutboundKey).value) Then
                  loRsChanges.Fields(RS_CL_mwrBatchLogOutboundKey).value = mnBatchKey
                  loRsChanges.Update
               End If
               
               RecordsWritten = RecordsWritten + 1
               
               If mStartChangeLogKey <= 0 Then
                  mStartChangeLogKey = loRsChanges.Fields(RS_CL_ID).value
               End If
               
               mEndChangeLogKey = loRsChanges.Fields(RS_CL_ID).value
               
               
               ' If we just processed a Schema record then we need to close the datagram and create a new one
               ' But only if there are records remaining to process
               ' AND the next record isn't another Schema record.
               
               If loRsChanges.Fields(RS_CL_ChangeType).value = "S" Then
                  loRsChanges.MoveNext
               
                  ' If we have any more records then close the schema datagram and create a new batch datagram
                  ' But not if the next record is another schema record, that will be handled up in the main loop
                  
                  If Not loRsChanges.EOF Then
                     If loRsChanges.Fields(RS_CL_ChangeType).value <> "S" Then
                        If CloseOldDatagram(BatchTypeKey, loRsSites!SiteID, loRsSites!ID, loRsSites!ShortSiteID) = False Then
                           CloseRecordset loRsChanges
                           CloseRecordset loRsSites
                           If Not moform Is Nothing Then
                              moform.TearDown
                           End If
                           CreateBatchDatagrams = False
                           goCon.RollbackTrans
                           Exit Function
                        End If
                        
'                        goSession.RaiseWorkflowStatus "......New StartChangeLogKey = " & loRsChanges.Fields(RS_CL_ID).value
                        goSession.RaiseWorkflowStatus "..."
                        
                        loRsOBW!StartChangeLogKey = loRsChanges.Fields(RS_CL_ID).value
                        loRsOBW.Update
                        goCon.CommitTrans
                        goCon.BeginTrans
                        
                        If StartNewDatagram(BatchTypeKey, loRsSites!SiteID, loRsSites!ID, loRsSites!ShortSiteID) = False Then
                           CloseRecordset loRsChanges
                           CloseRecordset loRsSites
                           If Not moform Is Nothing Then
                              moform.TearDown
                           End If
                           CreateBatchDatagrams = False
                           goCon.RollbackTrans
                           Exit Function
                        End If
                        
                     End If
                  End If
               Else
                  loRsChanges.MoveNext
               End If
            Loop
            '
            ' Create Integrity Trailer Record
            '
            goSession.RaiseWorkflowStatus "...Done looping through transactions, closing datagram now"
            
            If WriteTrailerRecord(BatchTypeKey, loRsSites!SiteID, loRsSites!ID) = False Then
               CloseRecordset loRsChanges
               CloseRecordset loRsSites
               CreateBatchDatagrams = False
               goCon.RollbackTrans
               Exit Function
            End If
            
            If Not moform Is Nothing Then
               If CloseDatagram(BatchTypeKey, loRsSites!SiteID, loRsSites!ID) = False Then
                  CloseRecordset loRsChanges
                  CloseRecordset loRsSites
                  If Not moform Is Nothing Then
                     moform.TearDown
                  End If
                  CreateBatchDatagrams = False
                  goCon.RollbackTrans
                  Exit Function
               End If
            End If
         Else
            If goSession.Site.SiteType = SITE_TYPE_SHORE Then
               goSession.RaiseWorkflowStatus "No transactions found to send to: " & loRsSites!SiteName
            Else
               goSession.RaiseWorkflowStatus "No transactions found to send to shore."
            End If
         End If
         CloseRecordset loRsChanges
         
         loRsOBW!IsSiteDone = True
         loRsOBW.Update
      goCon.CommitTrans
      goSession.RaiseWorkflowStatus "...Commit Transaction: Finished processing " & loRsSites!SiteName
      goSession.RaiseWorkflowStatus "..."
      
      loRsOBW.Requery
   goSession.MemoryLogIt "CreateBatchDatagrams 7 "
   Loop
   
   mStartChangeLogKey = 0
   mEndChangeLogKey = 0
   
   CloseRecordset loRsOBW
   CloseRecordset loRsChanges
   CloseRecordset loRsSites
      
   goSession.MemoryLogIt "CreateBatchDatagrams 8 "
      
   goCon.BeginTrans
      goSession.RaiseWorkflowStatus "...Begin Transaction: Looking for unsent datagrams."
      '
      ' Copy from WorkingFolder to ArchivedOK...
      '
      
      Set loRs = FetchOutboundBatchlogRS(BATCH_TYPE_SITE_DATAGRAM)
      Set fso = New FileSystemObject
      Do While Not loRs.EOF
         If Not IsNull(loRs!DatagramFormID) Then
            sArcFile = goSession.Workflow.ArchiveOkZipFile & "\" & loRs!DatagramFormID & ".dgm"
            If Not fso.FileExists(sArcFile) Then
               '
               ' Remake the file
               '
               If Not RemakeTransmissionFile(loRs!ID) Then
                  goCon.RollbackTrans
                  goSession.RaiseError "Error in mwReplicateBatchWork.CreateBatchDatagrams, File not found: " & sArcFile
                  goSession.RaiseWorkflowStatus "Operation Cancelled, RollbackTrans has been executed."
'                  CloseRecordset loRsChanges
'                  CloseRecordset loRsSites
                  If Not moform Is Nothing Then
                     moform.TearDown
                  End If
                  CreateBatchDatagrams = False
                  Exit Function
               End If
            End If
         End If
         loRs.MoveNext
      Loop
      goSession.MemoryLogIt "CreateBatchDatagrams 9 "
      
      goSession.RaiseWorkflowStatus "... Archiving datagrams and populating mwrAuditLog."
      goCon.Execute "DELETE FROM mwrOutboundWork"
      
      On Error GoTo ArchivingError
      '
      ' Archive transactions into mwrAuditLog/mwrAuditLogColumn
      '
      IsArchiving = True
      goSession.GUI.SetCursorWait "Archiving Transactions that have been transmitted"

      ArchiveChangeLog "DELETE FROM mwrAuditLogColumn WHERE mwrChangeLogKey IN (SELECT ID FROM mwrChangeLog WHERE mwrBatchLogOutboundKey Is Not Null)"
      ArchiveChangeLog "DELETE FROM mwrAuditLog WHERE ID IN (SELECT ID FROM mwrChangeLog WHERE mwrBatchLogOutboundKey Is Not Null)"
      
      ArchiveChangeLog "INSERT INTO mwrAuditLog SELECT * FROM mwrChangeLog WHERE mwrBatchLogOutboundKey Is Not Null"
      ArchiveChangeLog "INSERT INTO mwrAuditLogColumn SELECT * FROM mwrChangeLogColumn WHERE mwrChangeLogKey IN (SELECT ID FROM mwrAuditLog)"
      
      ArchiveChangeLog "DELETE FROM mwrChangeLogColumn WHERE mwrChangeLogKey IN (SELECT ID FROM mwrAuditLog)"
      ArchiveChangeLog "DELETE FROM mwrChangeLog WHERE ID IN (SELECT ID FROM mwrAuditLog)"
      
      'DELETE RECORDS FROM SHORE TO SHORE
      'Added By N.Angelakis On 03022009
      'DEV1568: Modified code to move items from mwrChangeLog to mwrAuditLog for Shore
      If goSession.Site.SiteType = SITE_TYPE_SHORE Then
        
         If goSession.ThisSite.IsLogIncomingTransactions Then
   
            ArchiveChangeLog "DELETE FROM mwrAuditLogColumn WHERE mwrChangeLogKey IN (SELECT ID FROM mwrChangeLog WHERE mwrBatchLogOutboundKey Is Null AND (IsLoggedTransaction = 1 or IsLoggedTransaction = -1) )"
            ArchiveChangeLog "DELETE FROM mwrAuditLog WHERE ID IN (SELECT ID FROM mwrChangeLog WHERE mwrBatchLogOutboundKey Is Null AND (IsLoggedTransaction = 1 or IsLoggedTransaction = -1))"
            
            ArchiveChangeLog "INSERT INTO mwrAuditLog SELECT * FROM mwrChangeLog WHERE mwrBatchLogOutboundKey Is Null AND (IsLoggedTransaction = 1 or IsLoggedTransaction = -1)"
            ArchiveChangeLog "INSERT INTO mwrAuditLogColumn SELECT * FROM mwrChangeLogColumn WHERE mwrChangeLogKey IN (SELECT ID FROM mwrAuditLog WHERE (IsLoggedTransaction = 1 or IsLoggedTransaction = -1))"
            
            ArchiveChangeLog "DELETE FROM mwrChangeLogColumn WHERE mwrChangeLogKey IN (SELECT ID FROM mwrAuditLog WHERE (IsLoggedTransaction = 1 or IsLoggedTransaction = -1))"
            ArchiveChangeLog "DELETE FROM mwrChangeLog WHERE ID IN (SELECT ID FROM mwrAuditLog WHERE (IsLoggedTransaction = 1 or IsLoggedTransaction = -1))"
         
         End If
        
        goSession.RaiseWorkflowStatus "...Archiving Shore Site specific Transactions."
        
        ArchiveChangeLog "DELETE FROM mwrAuditLogColumn WHERE mwrChangeLogKey IN (SELECT ID FROM mwrChangeLog WHERE mwcSitesKey = " & goSession.Site.SiteKey & " AND OriginatingSiteKey = " & goSession.Site.SiteKey & ")"
        ArchiveChangeLog "DELETE FROM mwrAuditLog WHERE ID IN (SELECT ID FROM mwrChangeLog WHERE mwcSitesKey = " & goSession.Site.SiteKey & " AND OriginatingSiteKey = " & goSession.Site.SiteKey & ")"
        
        ArchiveChangeLog "INSERT INTO mwrAuditLog SELECT * FROM mwrChangeLog WHERE mwcSitesKey=" & goSession.Site.SiteKey & " AND OriginatingSiteKey = " & goSession.Site.SiteKey
        ArchiveChangeLog "INSERT INTO mwrAuditLogColumn SELECT * FROM mwrChangeLogColumn WHERE mwrChangeLogKey IN (SELECT ID FROM mwrAuditLog)"
        
        ArchiveChangeLog "DELETE FROM mwrChangeLogColumn WHERE mwrChangeLogKey IN (SELECT ID FROM mwrAuditLog)"
        ArchiveChangeLog "DELETE FROM mwrChangeLog WHERE ID IN (SELECT ID FROM mwrAuditLog)"
      End If
      goSession.GUI.SetCursorNormal
      goSession.MemoryLogIt "CreateBatchDatagrams 10 "

      
      IsArchiving = False
ArchivingError:
'      If IsArchiving Then
'         goSession.RaiseError "Error in mwrReplicateBatchWork.CreateBatchDatagrams, Archiving from Change Logs to Audit Logs. " & vbCrLf & "Process will continue, manual cleanup is required." & vbCrLf & " Cmd: " & vbCrLf & vbCrLf & strArchivingCmd & vbCrLf & vbCrLf & "fails with error ", goCon.Errors(0).Number, goCon.Errors(0).Description
''         MsgBox "Process will continue, manual cleanup is required."
'      End If
      
   goCon.CommitTrans
   goSession.RaiseWorkflowStatus "...Preparation Complete, Transactions Committed. "
   If loRs.RecordCount < 1 Then
      goSession.RaiseWorkflowStatus "No outgoing transmission files were sent, process complete."
   Else
      goSession.RaiseWorkflowStatus "Successfully created outgoing transmission file(s), will now send them where SUSPEND is False..."
      goSession.MemoryLogIt "CreateBatchDatagrams 11 "
      
      loRs.MoveFirst
      Do While Not loRs.EOF
         Set moform = goSession.OpenNewForm
         If Not IsNull(loRs!DatagramFormID) Then
            moform.OpenForm (goSession.Workflow.ArchiveOkZipFile & "\" & loRs!DatagramFormID & ".dgm")
            If moform.IsOpen Then
               If goSession.Site.SiteType = SITE_TYPE_SHIP Then
                  sTargetSiteID = goSession.Site.TargetReplicateSiteID
                  'sEmailAddress = goSession.Site.GetShipProperty(sTargetSiteID, "WorkflowSendToAddress")
                  sEmailAddress = goSession.Site.WorkflowSendToAddress
               Else
                  sTargetSiteID = moform.FormProperties.TargetSiteID
                  sEmailAddress = goSession.Site.GetShipProperty(sTargetSiteID, "ShipWorkflowSendToAddress")
               End If
               ' Check Suspend
               If IsSuspendTransmit(sTargetSiteID) Then
                  goSession.RaiseWorkflowStatus "Suspend Transmission in effect, file not sent: " & loRs!DatagramFormID & _
                   ", to " & goSession.Site.GetSiteName(loRs!mwcSitesKey)
               Else
                  If moform.Send(sEmailAddress, , sTargetSiteID, NO_ARCHIVE) Then
                     If goSession.Site.SiteType = SITE_TYPE_SHORE Then
                        goSession.RaiseWorkflowStatus "Successfully sent batch datagram, " & loRs!DatagramFormID & _
                         ", to " & goSession.Site.GetSiteName(loRs!mwcSitesKey)
                     Else
                        goSession.RaiseWorkflowStatus "Successfully sent batch datagram, " & loRs!DatagramFormID & _
                         ", to Shore Workflow Agent."
                     End If
                     loRs.Fields("DateSent").value = Now()
                     loRs.Update
                     
                  Else
                     If goSession.Site.SiteType = SITE_TYPE_SHORE Then
                        goSession.RaiseWorkflowStatus "Error sending batch datagram, " & loRs!DatagramFormID & _
                         ", to " & goSession.Site.GetSiteName(loRs!mwcSitesKey)
                     Else
                        goSession.RaiseWorkflowStatus "Error sending batch datagram, " & loRs!DatagramFormID & _
                         ", to Shore Workflow Agent."
                     End If
                  End If
               End If
            End If
         End If
         moform.TearDown
         Set moform = Nothing
         loRs.MoveNext
      Loop
      goSession.RaiseWorkflowStatus "...completed Send Batch Datagram process."
   End If
   
   CloseRecordset loRs
   CreateBatchDatagrams = True
   goSession.MemoryLogIt "CreateBatchDatagrams exit "
   Exit Function
FunctionError:
   'Resume Next
   goSession.RaiseError "General Error in mwReplicateBatchWork.CreateBatchDatagrams. ", Err.Number, Err.Description
   Resume Next
   CloseRecordset loRsSites
   CloseRecordset loRsChanges
   CloseRecordset loRsOBW
   KillObject loAuditTrail
   CreateBatchDatagrams = False
   On Error Resume Next
   goCon.RollbackTrans
   goSession.MemoryLogIt "CreateBatchDatagrams err exit "
   
End Function
Private Sub ArchiveChangeLog(strArchivingCmd As String)
   On Error GoTo ArchivingError
   '
   ' Archive transactions into mwrAuditLog/mwrAuditLogColumn
   '
   goCon.Execute strArchivingCmd
   Exit Sub
ArchivingError:
   goSession.RaiseError "Error in mwrReplicateBatchWork.CreateBatchDatagrams, Archiving from Change Logs to Audit Logs. " & vbCrLf & "Process will continue, manual cleanup is required." & vbCrLf & " Cmd: " & vbCrLf & vbCrLf & strArchivingCmd & vbCrLf & vbCrLf & "fails with error ", goCon.Errors(0).Number, goCon.Errors(0).Description
End Sub
Private Function FetchOutboundWork() As Recordset
   Dim ssql As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   
   ssql = "SELECT * FROM mwrOutboundWork WHERE IsSiteDone = 0 ORDER BY FleetBatchKey, mwcSitesKey"
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open ssql, goCon, adOpenDynamic, adLockOptimistic
   Set FetchOutboundWork = loRs
   Set loRs = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.FetchOutboundWork. ", Err.Number, Err.Description
   Set FetchOutboundWork = Nothing
End Function
Private Function GetMinChangeLogKey() As Long
   Dim ssql As String
   Dim sAndQry As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   
   If goSession.ThisSite.IsLogIncomingTransactions Then
      sAndQry = " AND (IsLoggedTransaction Is Null or IsLoggedTransaction = 0) "
   End If
   
   ssql = "SELECT MIN(ID) FROM mwrChangeLog WHERE mwrBatchLogOutboundKey Is Null " & sAndQry

   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open ssql, goCon, adOpenDynamic, adLockOptimistic
   
   If loRs.RecordCount = 1 Then
      GetMinChangeLogKey = ZeroNull(loRs.Fields(0).value)
   Else
      GetMinChangeLogKey = -1
   End If
   
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.GetMinChangeLogKey. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function
Private Function GetMaxChangeLogKey() As Long
   Dim ssql As String
   Dim sAndQry As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   
   If goSession.ThisSite.IsLogIncomingTransactions Then
      sAndQry = " AND (IsLoggedTransaction Is Null or IsLoggedTransaction = 0) "
   End If
   
   ssql = "SELECT MAX(ID) FROM mwrChangeLog WHERE mwrBatchLogOutboundKey Is Null " & sAndQry

   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open ssql, goCon, adOpenDynamic, adLockOptimistic
   
   If loRs.RecordCount = 1 Then
      GetMaxChangeLogKey = ZeroNull(loRs.Fields(0).value)
   Else
      GetMaxChangeLogKey = -1
   End If
   
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.GetMaxChangeLogKey. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function

Private Function CloseOldDatagram(BatchTypeKey As Long, TargetSiteID As String, TargetSiteKey As Long, ShortSiteID As String) As Boolean
   On Error GoTo FunctionError
   
   If WriteTrailerRecord(BatchTypeKey, TargetSiteID, TargetSiteKey) = False Then
      CloseOldDatagram = False
      Exit Function
   End If
   
   If Not moform Is Nothing Then
      If CloseDatagram(BatchTypeKey, TargetSiteID, TargetSiteKey) = False Then
         CloseOldDatagram = False
         Exit Function
      End If
   End If
   
   mStartChangeLogKey = 0
   mEndChangeLogKey = 0
   
   CloseOldDatagram = True
   Exit Function
FunctionError:
   'Resume Next
   goSession.RaiseError "General Error in mwReplicateBatchWork.CloseOldDatagram. ", Err.Number, Err.Description
   CloseOldDatagram = False

End Function
Private Function StartNewDatagram(BatchTypeKey As Long, TargetSiteID As String, TargetSiteKey As Long, ShortSiteID As String) As Boolean
   On Error GoTo FunctionError
   
   mStartChangeLogKey = 0
   mEndChangeLogKey = 0
   
   If PrepareNewDatagram(BatchTypeKey, TargetSiteKey, ShortSiteID) = False Then
      StartNewDatagram = False
      Exit Function
   End If
   
   If WriteHeaderRecord(BatchTypeKey, TargetSiteID, TargetSiteKey) = False Then
      StartNewDatagram = False
      Exit Function
   End If
   StartNewDatagram = True
   Exit Function
FunctionError:
   'Resume Next
   goSession.RaiseError "General Error in mwReplicateBatchWork.StartNewDatagram. ", Err.Number, Err.Description
   StartNewDatagram = False

End Function

Private Function PrepareNewDatagram(BatchTypeKey As Long, TargetSiteID As Long, ShortSiteID As String) As Boolean
   Dim strTemplate As String
   Dim sFormID As String

   On Error GoTo FunctionError
   
   If goSession.Site.SiteType = SITE_TYPE_SHIP Then
      If Trim(goSession.Site.TargetReplicateSiteID) = "" Then
         goSession.RaiseError "Error in mwReplicateBatchWork.CreateBatchDatagrams, TargetSiteID missing for ship."
         PrepareNewDatagram = False
         Exit Function
      End If
      mnBatchKey = CreateOutboundBatchLog(BatchTypeKey, mnBatchNo, goSession.Site.GetSiteKey(goSession.Site.TargetReplicateSiteID))
   Else
      mnBatchKey = CreateOutboundBatchLog(BatchTypeKey, mnBatchNo, TargetSiteID)
   End If
   If mnBatchKey < 1 Then
      PrepareNewDatagram = False
      Exit Function
   End If
   ' New Form...
   strTemplate = FetchBatchKeyTemplate(BatchTypeKey)
   If strTemplate = "" Then
      goSession.RaiseError "Error in mwReplicateBatchWork.CreateBatchDatagrams, TemplateID missing from mwrBatchType: " & BatchTypeKey
      PrepareNewDatagram = False
      Exit Function
   End If
   If goSession.Site.SiteType = SITE_TYPE_SHIP Then
      sFormID = goSession.Site.ShortSiteID & goSession.Site.GetShipProperty(goSession.Site.TargetReplicateSiteID, "ShortSiteID") & _
                Format(mnBatchNo, "0000")
   Else
      sFormID = goSession.Site.ShortSiteID & ShortSiteID & Format(mnBatchNo, "0000")
   End If
   Set moform = goSession.OpenNewForm
   If Not moform.MakeNewForm(strTemplate, , , sFormID) Then
      PrepareNewDatagram = False
      Exit Function
   End If
   
   PrepareNewDatagram = True
   Exit Function
FunctionError:
   Resume Next
   goSession.RaiseError "General Error in mwReplicateBatchWork.PrepareNewDatagram. ", Err.Number, Err.Description
   PrepareNewDatagram = False
   
End Function
Private Function WriteHeaderRecord(BatchTypeKey As Long, TargetSiteID As String, TargetSiteKey As Long) As Boolean
   Dim sArcFile As String

   On Error GoTo FunctionError
   
   With moform.DataForm.Records(1).ConfigKeys
      .Add "RemakeTransmission", "False", "RemakeTransmission"
      If goSession.Site.SiteType <> SITE_TYPE_SHORE Then
         ' Ship to Shore
         .Add "TargetSiteID", CStr(goSession.Site.TargetReplicateSiteID), "TargetSiteID"
         .Add "TargetSiteKey", CStr(goSession.Site.GetSiteKey(goSession.Site.TargetReplicateSiteID)), "TargetSiteKey"
      Else
         ' Shore to Ship
         .Add "TargetSiteID", CStr(TargetSiteID), "TargetSiteID"
         .Add "TargetSiteKey", CStr(TargetSiteKey), "TargetSiteKey"
      End If

      .Add "BatchKey", CStr(mnBatchKey), "BatchKey"
      .Add "BatchType", CStr(BatchTypeKey), "BatchType"
      .Add "BatchNumber", CStr(mnBatchNo), "BatchNumber"

      ' added by TJM
      .Add "FleetBatchNumber", CStr(mNewFleetBatchKey), "FleetBatchNumber"

   End With
   
   WriteHeaderRecord = True
   Exit Function
FunctionError:
   'Resume Next
   goSession.RaiseError "General Error in mwReplicateBatchWork.WriteHeaderRecord. ", Err.Number, Err.Description
   WriteHeaderRecord = False
   
End Function
   
Private Function CloseDatagram(BatchTypeKey As Long, TargetSiteID As String, TargetSiteKey As Long) As Boolean
   Dim sArcFile As String
   
   DoEvents
   goSession.RaiseWorkflowStatus "...Writing datagram to WorkflowTemp"
   
   If Not moform.DataForm.WriteDataForm() Then
      goSession.RaiseError "Error in mwReplicateBatchWork.CreateBatchDatagrams, WriteDataForm Failed. Operation Cancelled."
      CloseDatagram = False
      Exit Function
   End If
   
   '
   ' Update the Batch Log entry
   '
   If FinalizeOutboundBatchlog(mnBatchKey, moform.FormProperties.FormID, _
     moform.DataForm.Records.Count, mNewFleetBatchKey) Then
      goSession.RaiseWorkflowStatus "...Finished Creating Batch Datagram: " & moform.FormProperties.FormID
   Else
      goSession.RaiseError "Error in mwReplicateBatchWork.CreateBatchDatagrams, datagram sent - finalize log entry failed !"
   End If
   sArcFile = goSession.Workflow.WorkingSessionFolder & "\" & moform.FormProperties.FormID & ".dgm"
   moform.TearDown

   DoEvents
   '
   ' Archive file now
   '
   goSession.RaiseWorkflowStatus "...Copying datagram to ArchivedOK"
   
   If sArcFile <> "" Then
      If Not goSession.Workflow.ArchiveWorkflowItem(sArcFile, ArchivedOK, True, SKIP_ERRORS) Then
         goSession.RaiseError "Error in mwReplicateBatchWork.CreateBatchDatagrams, Archive Failed on " & _
          sArcFile & ". Operation Cancelled."
         CloseDatagram = False
         Exit Function
      End If
   End If
   
   CloseDatagram = True
   
   Exit Function
FunctionError:
   'Resume Next
   goSession.RaiseError "General Error in mwReplicateBatchWork.CloseDatagram. ", Err.Number, Err.Description
   CloseDatagram = False
End Function
Private Function WriteTrailerRecord(BatchTypeKey As Long, TargetSiteID As String, TargetSiteKey As Long) As Boolean
   
   On Error GoTo FunctionError
   
   With moform.DataForm.AddRecord
      .Add "Operation", "I"
      .Add "COMMENT", "Batch Integrity Record"
      .Add "SiteID", goSession.Site.SiteID, "SiteID"
      If goSession.Site.SiteType <> SITE_TYPE_SHORE Then
         ' Ship to shore
         .Add "TargetSiteID", CStr(goSession.Site.TargetReplicateSiteID), "TargetSiteID"
         .Add "TargetSiteKey", CStr(goSession.Site.GetSiteKey(goSession.Site.TargetReplicateSiteID)), "TargetSiteKey"
      Else
         ' Shore to ship
         .Add "TargetSiteID", CStr(TargetSiteID), "TargetSiteID"
         .Add "TargetSiteKey", CStr(TargetSiteKey), "TargetSiteKey"
      End If
      .Add "BatchType", CStr(BatchTypeKey), "BatchType"
      .Add "BatchNumber", CStr(mnBatchNo), "BatchNumber"
      ' added by TJM
      .Add "FleetBatchNumber", CStr(mNewFleetBatchKey), "FleetBatchNumber"
      .Add "BatchKey", CStr(mnBatchKey), "BatchKey"
      .Add "TotalRecords", moform.DataForm.Records.Count
   End With
   WriteTrailerRecord = True
   
   Exit Function
FunctionError:
   'Resume Next
   goSession.RaiseError "General Error in mwReplicateBatchWork.WriteTrailerRecord. ", Err.Number, Err.Description
   WriteTrailerRecord = False
End Function

Public Function CreateSatelliteDatagram() As Boolean
   Dim ssql As String
   Dim loRsSites As Recordset
   Dim loRsChanges As Recordset
   Dim loRsColumns As Recordset
   Dim loRsDetail As Recordset
   Dim loRs As Recordset
   Dim loRec As ConfigKeys
   Dim loform As mwForm
   Dim loDbWork As mwDataWork
   Dim strTemplate As String
   Dim sTargetSiteID As String
   Dim IsError As Boolean
   Dim sEmailAddress As String
   Dim sFormID As String
   Dim s As String
   Dim sArcFile As String
   Dim sTargetFile As String
   Dim fso As FileSystemObject
   Dim nSiteType As Integer
   Dim RecordsWritten As Long
   Dim MaxRecordsForSite As Long
   Dim FormIsOpen As Boolean
   Dim TotalWrites As Long
   Dim IsSatellite As Boolean
   Dim sUserFolder As String
   Dim loWorkFlow As mwWorkflow
   Dim loReg As Registry
   Dim DefaultTarget As String
   
   On Error GoTo FunctionError
   
'   If goSession.Site.SiteType = SITE_TYPE_SHORE Then
'      goSession.RaiseError "CreateSatelliteDatagram is not valid on a SHORE site. Operation Cancelled."
'      CreateSatelliteDatagram = False
'      Exit Function
'   End If
   
   '
   ' Fetch Regional Settings
   '
   nSiteType = goSession.Site.SiteType
   goSession.Site.SiteType = SITE_TYPE_SHIP
   mDgGeneralDateFormat = DATAGRAM_DATE_FORMAT
   mDgShortDateFormat = DATAGRAM_SHORT_DATE_FORMAT
   mDgDecimalSeparator = DATAGRAM_DECIMAL_FORMAT
   
   If Len(goSession.ThisSite.SatelliteDBName) > 0 Then
      IsSatellite = True
   Else
      IsSatellite = False
   End If
   '
   ' Step 1 Fetch Sites to replicate to...
   '
   If IsSatellite = True Then
      goSession.RaiseWorkflowStatus "Creating Satellite Datagram to update the central database... This could take some time!"
   Else
      goSession.RaiseWorkflowStatus "Creating Restoration Datagram to restore the Ship database..."
   End If
   
   mOversizeBatchLastRecKey = 0
   
   goCon.BeginTrans
      '
      ' Fetch all changes...
      '
      Set loRsChanges = FetchSatelliteChangesRS
      DoEvents
      If loRsChanges Is Nothing Then
         goSession.RaiseError "Error in mwReplicateBatchWork.CreateSatelliteDatagram, unable to open changes table. Operation Cancelled."
         CloseRecordset loRsChanges
         CloseRecordset loRsSites
         CreateSatelliteDatagram = False
         goCon.RollbackTrans
         Exit Function
      End If
      '
      '
      ' Any Changes  ??
      '
      If loRsChanges.RecordCount > 0 Then
            
         
         ' New Form...
         strTemplate = "SYSTEM_CHANNEL"

         If IsSatellite = True Then
            MaxRecordsForSite = 0
         Else
            MaxRecordsForSite = 10000
         End If
         RecordsWritten = 0
         TotalWrites = 0
         
         '
         ' Loop through change records...
         '
         goSession.RaiseWorkflowStatus "...Writing " & loRsChanges.RecordCount & " Changes"
         
         Do While Not loRsChanges.EOF
               
            If TotalWrites > 0 And TotalWrites Mod 1000 = 0 Then
               goSession.RaiseWorkflowStatus "...Wrote " & TotalWrites
            End If
            
            If RecordsWritten = 0 Or (MaxRecordsForSite > 0 And RecordsWritten >= MaxRecordsForSite) Then
               If MaxRecordsForSite > 0 And RecordsWritten > 0 Then
                  goSession.RaiseWorkflowStatus "...Max Records Per Datagram reached, starting new datagram"
               End If
               
               If FormIsOpen = True Then
                  If Not loform.DataForm.WriteDataForm() Then
                     goSession.RaiseError "Error in mwReplicateBatchWork.CreateSatelliteDatagram, WriteDataForm Failed. Operation Cancelled."
                     CloseRecordset loRsChanges
                     goCon.RollbackTrans
                     loform.TearDown
                     CreateSatelliteDatagram = False
                     Exit Function
                  End If
                  FormIsOpen = False
                  
                  '
                  ' Update the Batch Log entry
                  '
                  If FinalizeOutboundBatchlog(mnBatchKey, loform.FormProperties.FormID, _
                    loform.DataForm.Records.Count, mNewFleetBatchKey) Then
                     
                     If IsSatellite = True Then
                        goSession.RaiseWorkflowStatus "...Finished Creating Satellite Datagram: " & loform.FormProperties.FormID
                     Else
                        goSession.RaiseWorkflowStatus "...Created Restoration Datagram: " & loform.FormProperties.FormID
                     End If
                  Else
                     goSession.RaiseError "Error in mwReplicateBatchWork.CreateSatelliteDatagram, datagram sent - finalize log entry failed !"
                  End If
                  loform.TearDown
                  
               End If
               
               mnBatchKey = CreateOutboundBatchLog(1, mnBatchNo)
               
               If Len(goSession.ThisSite.SatelliteDBName) > 0 Then
                  sFormID = goSession.ThisSite.SatelliteDBName & "_" & Format(Now(), "YYMMDDhhmmss")
               Else
                  sFormID = goSession.Site.SiteID & "_" & Format(Now(), "YYMMDDhhmmss")
               End If
               
               Set loform = goSession.OpenNewForm
               If Not loform.MakeNewForm(strTemplate, , , sFormID) Then
                  CloseRecordset loRsChanges
                  CreateSatelliteDatagram = False
                  goCon.RollbackTrans
                  Exit Function
               End If
               '
               ' Populate Target Info in Header...
               '
               If Len(goSession.ThisSite.SatelliteDBName) > 0 Then
                  With loform.DataForm.Records(1).ConfigKeys
                     .Add "SatelliteDBName", goSession.ThisSite.SatelliteDBName
                  End With
               End If
         
               FormIsOpen = True
               RecordsWritten = 0
            End If
            
            DoEvents
            Select Case loRsChanges.Fields(RS_CL_ChangeType).value
               Case "D"
                  ' Delete Transaction
                  With loform.DataForm.AddRecord
                     .Add "OPERATION", "D", "OPERATION"
                     .Add "TABLE", loRsChanges.Fields(RS_CT_Tablename).value, "TABLE"
                     .Add "ID", loRsChanges.Fields(RS_CL_RecordKey).value, "ID"
                  End With
               Case "A"
                  ' Add Transaction
                  Set loRec = loform.DataForm.AddRecord
                  loRec.Add "OPERATION", "A", "OPERATION"
                  loRec.Add "TABLE", loRsChanges.Fields(RS_CT_Tablename).value, "TABLE"
                  If Not WriteChangeColumns(loRec, loRsChanges.Fields(RS_CL_ID).value) Then
                     CloseRecordset loRsChanges
                     goSession.RaiseWorkflowStatus "Error Creating Batch Datagrams. Process cancelled."
                     goCon.RollbackTrans
                     Exit Function
                  End If
               Case "M"
                  ' Modify Transaction
                  Set loRec = loform.DataForm.AddRecord
                  loRec.Add "OPERATION", "M", "OPERATION"
                  loRec.Add "TABLE", loRsChanges.Fields(RS_CT_Tablename).value, "TABLE"
                  loRec.Add "ID", loRsChanges.Fields(RS_CL_RecordKey).value, "ID"
                  If Not WriteChangeColumns(loRec, loRsChanges.Fields(RS_CL_ID).value) Then
                     CloseRecordset loRsChanges
                     goSession.RaiseWorkflowStatus "Error Creating Batch Datagrams. Process cancelled."
                     goCon.RollbackTrans
                     Exit Function
                  End If
                  
               Case "S"
                  ' Schema Transaction
                  Set loRec = loform.DataForm.AddRecord
                  loRec.Add "OPERATION", "S", "OPERATION"
                  loRec.Add "TABLE", loRsChanges.Fields(RS_CT_Tablename).value, "TABLE"
                  If Not WriteChangeColumns(loRec, loRsChanges.Fields(RS_CL_ID).value) Then
                     CloseRecordset loRsChanges
                     goSession.RaiseWorkflowStatus "Error Creating Batch Datagrams. Process cancelled."
                     goCon.RollbackTrans
                     Exit Function
                  End If
                  
               Case Else
                  goSession.RaiseWorkflowStatus "Warning in mwReplicateBatchWork. invalid change type code. Key: " & loRsChanges.Fields(RS_CL_ID).value
            End Select
            RecordsWritten = RecordsWritten + 1
            TotalWrites = TotalWrites + 1
            
            '
            ' Update BatchKey in Change Log record if it isn't already set to the Fleet Batch Key
            '
            If IsNull(loRsChanges.Fields(RS_CL_mwrBatchLogOutboundKey).value) Then
               loRsChanges.Fields(RS_CL_mwrBatchLogOutboundKey).value = mnBatchKey
               loRsChanges.Update
            End If
            loRsChanges.MoveNext
         Loop
         
         If FormIsOpen = True Then
            If IsSatellite = True Then
               goSession.RaiseWorkflowStatus "Writing Satellite Datagram to disk."
            Else
               goSession.RaiseWorkflowStatus "Writing Restoration Datagram to disk."
            End If
            DoEvents
            If Not loform.DataForm.WriteDataForm() Then
               goSession.RaiseError "Error in mwReplicateBatchWork.CreateSatelliteDatagram, WriteDataForm Failed. Operation Cancelled."
               CloseRecordset loRsChanges
               goCon.RollbackTrans
               loform.TearDown
               CreateSatelliteDatagram = False
               Exit Function
            End If
            FormIsOpen = False
            
            '
            ' Update the Batch Log entry
            '
            If FinalizeOutboundBatchlog(mnBatchKey, loform.FormProperties.FormID, _
              loform.DataForm.Records.Count, mNewFleetBatchKey) Then
               
               If IsSatellite = True Then
                  goSession.RaiseWorkflowStatus "...Created Satellite Datagram: " & loform.FormProperties.FormID
               Else
                  goSession.RaiseWorkflowStatus "...Created Restoration Datagram: " & loform.FormProperties.FormID
               End If
               DoEvents
            Else
               goSession.RaiseError "Error in mwReplicateBatchWork.CreateSatelliteDatagram, datagram sent - finalize log entry failed !"
            End If
            loform.TearDown
            
         End If
            
         DoEvents

      Else
         goSession.RaiseWorkflowStatus "No transactions found to send to the Central Ship database."
      End If
      DoEvents
      CloseRecordset loRsChanges
      

      '
      ' Copy from WorkingFolder to ArchivedOK...
      '

      Set fso = New FileSystemObject
      Set loRs = New Recordset
      loRs.CursorLocation = adUseClient
      ssql = "select * from mwrBatchLogOutbound WHERE DateSent IS NULL and mwrBatchTypeKey=1"
      loRs.Open ssql, goCon, adOpenDynamic, adLockOptimistic
      If loRs.RecordCount < 1 Then
         goSession.RaiseWorkflowStatus "No Outbound transactions to send."
      Else
         ' prompt for UserFolder to copy to
         
         Set loReg = New Registry
         loReg.BaseRegistry = BASE_REGISTRY & "CreateSatelliteDatagram"
      
         DefaultTarget = loReg.GetRegStr("DefaultTargetLocation", "")
         
         Set loWorkFlow = New mwWorkflow
'         sUserFolder = loWorkFlow.GetTargetFolder(DefaultTarget, "Where would you like to put the datagram?" & vbCrLf & vbCrLf & _
'                                                    "Select the Drive and Folder where you want to put the datagram. A copy will also be left in the Transport\ArchivedOk folder.")
         sUserFolder = goSession.Workflow.SendByMediaFolder
         Set loWorkFlow = Nothing
         goSession.RaiseWorkflowStatus "Datagram will be placed in the SendByMedia folder: " & sUserFolder
'         MsgBox "Datagram will be placed in the SendByMedia folder: " & sUserFolder, vbOKOnly, "Satellite Transmission File Location"
         DoEvents
      
'         If sUserFolder <> "" Then
'            loReg.SetReg "DefaultTargetLocation", sUserFolder
'         Else
'            sUserFolder = goSession.Workflow.SendByCdromFolder
'            goSession.RaiseWorkflowStatus "No Target Location selected, datagrams will be placed in the SendByCDROM folder: " & sUserFolder
'            MsgBox "No Target Location selected, datagrams will be placed in the SendByCDROM folder: " & sUserFolder, vbOKOnly, "No target folder selected"
'            If sUserFolder = "" Then
'               goSession.RaiseWorkflowStatus "The Workflow.SendByCDROM seems to be NULL. Datagram will be found in the ArchivedOK folder"
'               MsgBox "The Workflow.SendByCDROM seems to be NULL. Datagram will be found in the ArchivedOK folder", vbOKOnly, "SendByCDROM folder is NULL"
'            End If
'         End If
         KillObject loReg
         
         Do While Not loRs.EOF
            If Not IsNull(loRs!DatagramFormID) Then
               sArcFile = goSession.Workflow.WorkingSessionFolder & "\" & loRs!DatagramFormID & ".dgm"
               sTargetFile = goSession.Workflow.SendByMediaFolder & "\" & loRs!DatagramFormID & ".dgm"
               fso.CopyFile sArcFile, sTargetFile
               
               If sUserFolder <> "" And fso.FolderExists(sUserFolder) Then
                  fso.CopyFile sArcFile, sUserFolder & "\" & loRs!DatagramFormID & ".dgm"
                  goSession.RaiseWorkflowStatus "Datagram has been copied to the Target Location: " & sUserFolder & "\" & loRs!DatagramFormID & ".dgm"
                  DoEvents
'               Else
'                  goSession.RaiseWorkflowStatus "No Target Location selected, datagram has been placed in the SendByCDROM folder: " & sTargetFile
'                  sUserFolder = goSession.Workflow.SendByCdromFolder
               End If
   
               If Not goSession.Workflow.ArchiveWorkflowItem(sArcFile, ArchivedOK, True, SKIP_ERRORS) Then
                  sArcFile = goSession.Workflow.ArchiveOkZipFile & "\" & loRs!DatagramFormID & ".dgm"
                  If Not fso.FileExists(sArcFile) Then
                     goSession.RaiseError "Error in mwReplicateBatchWork.CreateSatelliteDatagram, File not found: " & sArcFile
                     goSession.RaiseWorkflowStatus "Operation Cancelled, RollbackTrans has been executed."
                     CloseRecordset loRsChanges
                     CloseRecordset loRs
                     goCon.RollbackTrans
                     If Not loform Is Nothing Then
                        loform.TearDown
                     End If
                     KillObject fso
                     CreateSatelliteDatagram = False
                     Exit Function
                  End If
               End If
               loRs!DateSent = Now()
               loRs.Update
            End If
            loRs.MoveNext
         Loop
         
'         DirAttachmentExport sUserFolder
      End If
      
      goSession.RaiseWorkflowStatus "Saving changes to the database, This could take some time!"
      DoEvents
      
      goCon.Execute "INSERT INTO mwrAuditLog SELECT * FROM mwrChangeLog WHERE mwrBatchLogOutboundKey Is Not Null"
      goCon.Execute "INSERT INTO mwrAuditLogColumn SELECT * FROM mwrChangeLogColumn WHERE mwrChangeLogKey IN (SELECT ID FROM mwrAuditLog)"
      goCon.Execute "DELETE FROM mwrChangeLogColumn WHERE ID IN (SELECT ID FROM mwrAuditLogColumn)"
      goCon.Execute "DELETE FROM mwrChangeLog WHERE ID IN (SELECT ID FROM mwrAuditLog)"

   goCon.CommitTrans
   goSession.RaiseWorkflowStatus "Create Satellite Datagram process completed."
   DoEvents
   goSession.Site.SiteType = nSiteType
   
   CreateSatelliteDatagram = True
   CloseRecordset loRs
   KillObject fso
   Exit Function
FunctionError:
   'Resume Next
   goSession.RaiseError "General Error in mwReplicateBatchWork.CreateSatelliteDatagram. " & vbCrLf & "Operation Cancelled, RollbackTrans has been executed.", Err.Number, Err.Description
   CloseRecordset loRsChanges
   KillObject fso
   KillObject loReg
   CreateSatelliteDatagram = False
   On Error Resume Next
   goCon.RollbackTrans
End Function
Private Function DirAttachmentExport(sTargetFolder As String) As Boolean
   
   Dim loRsFAB As Recordset
   Dim loRsDirSite As Recordset
   Dim loRsDirAtt As Recordset
   Dim ssql As String
   Dim sSQL2 As String
   Dim nBatchKey As Long
   Dim nBatchType As Long
   Dim nBatchNo As Long
   Dim sFileName As String
   Dim sFullFilename As String
   Dim loRepWork As mwReplicateBatchWork
   Dim i As Integer
   Dim sEmail As String
   Dim SendByMedia As Boolean
   Dim CloseDAB As Boolean
   Dim LastSite As Long
   
   On Error GoTo FunctionError
   
   Set loRepWork = New mwReplicateBatchWork
   ssql = "select * FROM mwEventDistributeSite " & _
            " WHERE IsDistribute = -1 and mwcSitesKey <>" & goSession.Site.SiteKey & _
            " AND mwrBatchLogOutboundKey Is Null " & _
            " AND mwEventDetailKey IN (SELECT smDirectiveKey from smDirAttachment) " & _
            " ORDER BY mwcSitesKey, mwEventDetailKey"
   
   Set loRsDirSite = New Recordset
   loRsDirSite.CursorLocation = adUseClient
   loRsDirSite.Open ssql, goCon, adOpenDynamic, adLockOptimistic
   
   sSQL2 = "select * FROM smDirAttachment ORDER BY ID"
   Set loRsDirAtt = New Recordset
   loRsDirAtt.CursorLocation = adUseClient
   loRsDirAtt.Open sSQL2, goCon, adOpenDynamic, adLockOptimistic
   
   '
   ' for each ship...
   '
   CloseDAB = False
   LastSite = 0
   
   Do While Not loRsDirSite.EOF
   
      loRsDirAtt.Filter = "smDirectiveKey = " & loRsDirSite!mwEventDetailKey
      
      If loRsDirAtt.RecordCount > 0 Then
         ' have some to send
         '
         
         goCon.BeginTrans
         ' Make RS to persist and send
         '
         If LastSite <> loRsDirSite!mwcSitesKey Then
         
            LastSite = loRsDirSite!mwcSitesKey
            
            SendByMedia = True
         
            Set loRsFAB = New Recordset
            With loRsFAB.Fields
               .Append "ID", adInteger, 4
               .Append "mwcSitesKey", adInteger, 4
               .Append "BatchType", adInteger, 4
               .Append "BatchKey", adInteger, 4
               .Append "BatchNumber", adInteger, 4
               .Append "FileBlob", adLongVarBinary, LONG_BIN_DEFINED_SIZE, adFldLong
            End With
            loRsFAB.Open
         '
            If SendByMedia Then
               nBatchType = BT_DA_BLOB_OUTBOUND_MEDIA
            Else
               nBatchType = BT_DA_BLOB_OUTBOUND_MODEM
            End If
            
            nBatchKey = loRepWork.CreateOutboundBatchLog(nBatchType, nBatchNo, loRsDirSite!mwcSitesKey)
            If nBatchKey < 1 Then
               goCon.RollbackTrans
               goSession.RaiseWorkflowStatus "Cancelling Directive Attachments Replication Due to an error."
               CloseRecordset loRsDirSite
               CloseRecordset loRsDirAtt
               DirAttachmentExport = False
               'SettingCheckbox = False
               Exit Function
            End If
         End If
         
         loRsDirAtt.MoveFirst
         
         Do While Not loRsDirAtt.EOF
            loRsFAB.AddNew
            loRsFAB!ID = loRsDirAtt!ID
            loRsFAB!mwcSitesKey = loRsDirSite!mwcSitesKey
            loRsFAB!BatchType = nBatchType
            loRsFAB!BatchKey = nBatchKey
            loRsFAB!BatchNumber = nBatchNo
            loRsFAB!FileBlob = loRsDirAtt!FileBlob
            loRsFAB.Update
            
            loRsDirAtt.Update
            loRsDirAtt.MoveNext
            
            loRsDirSite!mwrBatchLogOutboundKey = nBatchKey
            loRsDirSite.Update
         Loop
      End If
      loRsDirSite.MoveNext
      
      If loRsDirSite.EOF = True Then
         CloseDAB = True
      ElseIf LastSite <> loRsDirSite!mwcSitesKey Then
         CloseDAB = True
      Else
         CloseDAB = False
      End If
      
      If CloseDAB = True Then
         '
         ' Save/Archive File
         ' Update Outboundlog
         ' Commit Transaction
         '
         sFileName = "da" & goSession.Site.ShortSiteID & goSession.User.GetNextFormSequenceNumber
         sFullFilename = goSession.Workflow.ArchiveOkZipFile & "\" & sFileName & ".DAB"
         loRsFAB.Save sFullFilename, adPersistADTG
         
         If Not loRepWork.FinalizeOutboundBatchlog(nBatchKey, sFileName, loRsFAB.RecordCount) Then
            goCon.RollbackTrans
            CloseRecordset loRsDirSite
            CloseRecordset loRsDirAtt
            DirAttachmentExport = False
            'SettingCheckbox = False
            Exit Function
         End If
         
         goCon.CommitTrans
         '
         ' Try and Send it...
         '
         If SendByMedia Then
            If Not goSession.Workflow.SendByMedia(sFullFilename, ZIP_IT, , , , CDROM_IT, LastSite) Then
               goSession.RaiseWorkflowStatus "Error in mwReplicateBatchWork.DirAttachmentExport. File was prepared but not sent for site: " & LastSite
            Else
               goSession.RaiseWorkflowStatus "Directive Attachments Transmission " & str(nBatchNo) & " successfully sent to : " & LastSite
            End If
         Else
            sEmail = goSession.Site.GetShipProperty("", "ShipWorkflowSendToAddress", LastSite)
            
            
            If Not goSession.Workflow.SendByEmail(sFullFilename, "Directive Attachments Transmission " & str(nBatchNo) & " From " & goSession.Site.SiteName, sEmail, , LastSite) Then
               goSession.RaiseWorkflowStatus "Error in mwReplicateBatchWork.DirAttachmentExport. File was prepared but not sent for site: " & LastSite
            Else
               goSession.RaiseWorkflowStatus "Directive Attachments Transmission successfully " & str(nBatchNo) & " sent to : " & LastSite
            End If
         End If
         '
         ' cleanup...
         '
         CloseRecordset loRsFAB
      End If
   Loop
   CloseRecordset loRsDirSite
   CloseRecordset loRsDirAtt
   DirAttachmentExport = True
   If SendByMedia Then
      MsgBox "Directive Attachments Transmission Files for CDROM Transmission may be found in: " & _
             vbCrLf & sFullFilename  '  goSession.Workflow.GetExtendedProperty("SendByCdromFolder")
   End If
   Exit Function
   
FunctionError:
   goSession.RaisePublicError "General Error in mwReplicateBatchWork.DirAttachmentExport. ", Err.Number, Err.Description
   goCon.RollbackTrans
   CloseRecordset loRsDirSite
   CloseRecordset loRsDirAtt
   DirAttachmentExport = False
End Function


Private Function AddRecordField(ByRef loRec As ConfigKeys, FieldName As String, FieldValue As String, _
 FieldType As Long)
   Dim strTemp As String
   On Error GoTo FunctionError
   Select Case FieldType
      Case Is = adInteger        ' (3)
         loRec.Add FieldName, FieldValue, FieldName
      Case Is = adDouble         ' (5)
         '
         ' Assume Changelog has been updated correctly with "." decimal point value...
         '
      
         'strTemp = Format(FieldValue, "#########0.0##")
         ' Euro Symbol crap...
'         If mDgDecimalSeparator = "," Then
'            strTemp = Replace(FieldValue, ".", "")
'            strTemp = Replace(strTemp, ",", ".")
'         Else
'            strTemp = Replace(strTemp, ",", ".")
'         End If
         
         loRec.Add FieldName, FieldValue, FieldName
      Case Is = adDate, adDBTimeStamp           ' (7) (135)
         If Format(FieldValue, DATAGRAM_DATE_FORMAT) > "1899-12-30 00:00:00" Then
            'loRec.Add FieldName, Format(FieldValue, mDgGeneralDateFormat), FieldName
            loRec.Add FieldName, MakeDatagramDate(FieldValue), FieldName
         End If
      Case Is = adVarWChar       ' (202) Text
         strTemp = FixString(FieldValue)
         loRec.Add FieldName, strTemp, FieldName
      Case Is = adLongVarWChar   ' (203) - Memo Field
         strTemp = FixString(FieldValue)
         loRec.Add FieldName, strTemp, FieldName
      Case Is = adBoolean   ' (11) - Memo Field
         If UCase(FieldValue) = "TRUE" Then
            loRec.Add FieldName, Replace("1", vbCrLf, "|"), FieldName
         ElseIf UCase(FieldValue) = "FALSE" Then
            loRec.Add FieldName, Replace("0", vbCrLf, "|"), FieldName
         Else
            loRec.Add FieldName, Replace(FieldValue, vbCrLf, "|"), FieldName
         End If
      Case Else
         strTemp = FixString(FieldValue)
         loRec.Add FieldName, strTemp, FieldName
   End Select
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in adDBTimeStamp mwReplicateBatchWork.AddRecordField. ", Err.Number, Err.Description
End Function
Private Function FixString(strVal As String) As String
   Dim strTemp As String
   Dim ByteArray() As Byte
   Dim xx As Long
   
   ' Look for and fix UniCode characters.
   
   On Error GoTo FunctionError
   
   strTemp = Replace(strVal, vbCrLf, "|")
   strTemp = Replace(strTemp, vbLf, "|")
   
   ByteArray = strTemp
   
   For xx = 1 To UBound(ByteArray) Step 2
      If ByteArray(xx) <> 0 Then
         ByteArray(xx - 1) = 32     ' Replace uniCode character with a blank
         ByteArray(xx) = 0
      End If
   Next
   
   FixString = ByteArray
   
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwReplicateBatchWork.FixString.", Err.Number, Err.Description
   FixString = strVal
End Function

Private Function WriteChangeColumns(oRec As ConfigKeys, ChangeLogKey As Long) As Boolean
   Dim ssql As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   ssql = "SELECT *  from mwrChangeLogColumn where mwrChangeLogKey=" & ChangeLogKey & " ORDER BY ID"
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open ssql, goCon, adOpenForwardOnly, adLockReadOnly
   Do While Not loRs.EOF
      If Not IsNull(loRs!NewValue) Then
         AddRecordField oRec, loRs!ColumnName, loRs!NewValue, loRs!ColumnDataType
      ElseIf Not IsNull(loRs!NewValueLarge) Then
         AddRecordField oRec, loRs!ColumnName, loRs!NewValueLarge, loRs!ColumnDataType
      Else
         goSession.RaiseWorkflowStatus "Warning in mwReplicateBatchWork.WriteChangeColumns, no value to replicate. Change Log/Column Keys/Column Name: " & _
          ChangeLogKey & "/" & loRs!ID & "/" & loRs!ColumnName
      End If
      loRs.MoveNext
   Loop
   loRs.Close
   Set loRs = Nothing
   WriteChangeColumns = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwReplicateBatchWork.WriteChangeColumns. ", Err.Number, Err.Description
End Function

Private Function WriteAuditChangeColumns(oRec As ConfigKeys, ChangeLogKey As Long) As Boolean
   Dim ssql As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   ssql = "SELECT *  from mwrAuditLogColumn where mwrChangeLogKey=" & ChangeLogKey & " ORDER BY ID"
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open ssql, goCon, adOpenForwardOnly, adLockReadOnly
   Do While Not loRs.EOF
      If Not IsNull(loRs!NewValue) Then
         AddRecordField oRec, loRs!ColumnName, loRs!NewValue, loRs!ColumnDataType
      ElseIf Not IsNull(loRs!NewValueLarge) Then
         AddRecordField oRec, loRs!ColumnName, loRs!NewValueLarge, loRs!ColumnDataType
      Else
         goSession.RaiseWorkflowStatus "Warning in mwReplicateBatchWork.WriteAuditChangeColumns, no value to replicate. Key: " & loRs!ID
      End If
      loRs.MoveNext
   Loop
   loRs.Close
   Set loRs = Nothing
   WriteAuditChangeColumns = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwReplicateBatchWork.WriteAuditChangeColumns. ", Err.Number, Err.Description
End Function


Public Function MakeDatagramDate(RawDate As String, Optional IsShortDate As Boolean) As String
   Static IsErrored As Boolean
   Dim s As String
   On Error GoTo FunctionError
   
   s = Format(RawDate, DATAGRAM_DATE_FORMAT)
   
   MakeDatagramDate = s
   Exit Function
FunctionError:
   If Not IsErrored Then
      goSession.RaiseError "General Error in mwReplicateBatchWork.MakeDatagramDate. ", Err.Number, Err.Description
      IsErrored = True
   End If
   MakeDatagramDate = RawDate
End Function



Friend Function ValidateIncomingDgmBatch(oHRec As ConfigKeys, oBiRec As ConfigKeys, TotalRecords As Long) As Boolean
   Dim loRs As Recordset
   Dim nSourceSiteKey As Long
   Dim nBatchType As Long
   Dim nBatchNumber As Long
   Dim nExpBatchNumber As Long
   Dim sTargetSiteID As String
   Dim nTargetSiteKey As Long
   Dim ssql As String
   On Error GoTo FunctionError
   '
   ' Batch Type First
   '
   If IsNumeric(oHRec.GetKeyValue("BATCHTYPE")) Then
      nBatchType = oHRec.GetKeyValue("BATCHTYPE")
   Else
      goSession.RaiseWorkflowStatus "Error in mwReplicateBatchWork.ValidateIncomingDgmBatch, BATCHTYPE is missing."
      ValidateIncomingDgmBatch = False
      Exit Function
   End If
   
   '
   ' Batch Number First
   '
   If IsNumeric(oHRec.GetKeyValue("BATCHNUMBER")) Then
      nBatchNumber = oHRec.GetKeyValue("BATCHNUMBER")
   Else
      goSession.RaiseWorkflowStatus "Error in mwReplicateBatchWork.ValidateIncomingDgmBatch, BATCHNUMBER is missing."
      ValidateIncomingDgmBatch = False
      Exit Function
   End If
   
   '
   ' Target Site...
   '
   ' 11/2005 - Switch to use SiteKey and NOT SiteID (which can change) for validation...ms
   '
   If IsNumeric(oHRec.GetKeyValue("TARGETSITEKEY")) Then
      nTargetSiteKey = oHRec.GetKeyValue("TARGETSITEKEY")
      sTargetSiteID = goSession.Site.GetSiteID(nTargetSiteKey)
      nSourceSiteKey = oHRec.GetKeyValue("SiteKey")
   Else
      sTargetSiteID = oHRec.GetKeyValue("TARGETSITEID")
      nTargetSiteKey = goSession.Site.GetSiteKey(sTargetSiteID)
      nSourceSiteKey = goSession.Site.GetSiteKey(oHRec.GetKeyValue("SiteID"))
   End If
   If Trim(UCase(sTargetSiteID)) <> Trim(UCase(goSession.Site.SiteID)) Then
      goSession.RaiseWorkflowStatus "Error in mwReplicateBatchWork.ValidateIncomingDgmBatch, Target Site is incorrect: " & sTargetSiteID
      ValidateIncomingDgmBatch = False
      Exit Function
   End If
   '
   ' Site Key valid (where batch is from)...
   '
   If nSourceSiteKey < 1 Then
      goSession.RaiseWorkflowStatus "Error in mwReplicateBatchWork.ValidateIncomingDgmBatch, SiteID is missing or incorrect: " & oHRec.GetKeyValue("SiteID")
      ValidateIncomingDgmBatch = False
      Exit Function
   End If
   
   goSession.RaiseWorkflowStatus "Validating Batch Transmission Type/Number: " & _
     nBatchType & "/" & nBatchNumber & " received from: " & goSession.Site.GetSiteName(nSourceSiteKey)
   
   '
   ' Valid Batch Sequence Number ?
   '
   If Not ValidateIncomingSequence(nSourceSiteKey, nBatchType, nBatchNumber, nExpBatchNumber) Then
      '
      ' Request Alert here... ?
      '
      
      ValidateIncomingDgmBatch = False
      Exit Function
   End If
   '
   ' Batch Integrity Check
   '
   If oHRec.GetKeyValue("TARGETSITEID") <> oBiRec.GetKeyValue("TARGETSITEID") Then
      goSession.RaiseWorkflowStatus "Error in mwReplicateBatchWork.ValidateIncomingDgmBatch, Batch Integrity Record Check Failed. Header TARGETSITEID doesn't match Integrity record"
      ValidateIncomingDgmBatch = False
      Exit Function
   ElseIf oHRec.GetKeyValue("BATCHTYPE") <> oBiRec.GetKeyValue("BATCHTYPE") Then
      goSession.RaiseWorkflowStatus "Error in mwReplicateBatchWork.ValidateIncomingDgmBatch, Batch Integrity Record Check Failed. Header BATCHTYPE doesn't match Integrity record"
      ValidateIncomingDgmBatch = False
      Exit Function
   ElseIf oHRec.GetKeyValue("BATCHNUMBER") <> oBiRec.GetKeyValue("BATCHNUMBER") Then
      goSession.RaiseWorkflowStatus "Error in mwReplicateBatchWork.ValidateIncomingDgmBatch, Batch Integrity Record Check Failed. Header BATCHNUMBER doesn't match Integrity record"
      ValidateIncomingDgmBatch = False
      Exit Function
   ElseIf oHRec.GetKeyValue("SITEID") <> oBiRec.GetKeyValue("SITEID") Then
      goSession.RaiseWorkflowStatus "Error in mwReplicateBatchWork.ValidateIncomingDgmBatch, Batch Integrity Record Check Failed. Header SITEID doesn't match Integrity record"
      ValidateIncomingDgmBatch = False
      Exit Function
   ElseIf Not IsNumeric(oBiRec.GetKeyValue("TOTALRECORDS")) Then
      goSession.RaiseWorkflowStatus "Error in mwReplicateBatchWork.ValidateIncomingDgmBatch, Batch Integrity Record Check Failed. TOTALRECORDS is not numeric"
      ValidateIncomingDgmBatch = False
      Exit Function
   ElseIf CLng(oBiRec.GetKeyValue("TOTALRECORDS")) <> TotalRecords Then
      goSession.RaiseWorkflowStatus "Error in mwReplicateBatchWork.ValidateIncomingDgmBatch, Batch Integrity Record Check Failed. TOTALRECORDS = " & oBiRec.GetKeyValue("TOTALRECORDS") & ", actual TotalRecords = " & TotalRecords
      ValidateIncomingDgmBatch = False
      Exit Function
   End If
   '
   ' Save OutboundBatchKey
   If Not IsNumeric(oBiRec.GetKeyValue("BATCHKEY")) Then
      goSession.RaiseWorkflowStatus "Error in mwReplicateBatchWork.ValidateIncomingDgmBatch, BatchKey missing in Integrity Record."
      ValidateIncomingDgmBatch = False
      Exit Function
   End If
   
   mIncomingOutboundBatchKey = CLng(oBiRec.GetKeyValue("BATCHKEY"))
   '
   ' Falls through, We Are Cool...
   '
   goSession.RaiseWorkflowStatus "mwReplicateBatchWork.ValidateIncomingDgmBatch, Batch Integrity passed."
   ValidateIncomingDgmBatch = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwReplicateBatchWork.ValidateIncomingDgmBatch. ", Err.Number, Err.Description
   ValidateIncomingDgmBatch = False
End Function


Public Function MaybeRequestResend(SiteKey As Long, BatchTypeKey As Long, ReceivedBatchNumber As Long, _
  ExpectedBatchNumber As Long) As Boolean
   On Error GoTo FunctionError
   


   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwReplicateBatchWork.MaybeRequestResend. ", Err.Number, Err.Description
   MaybeRequestResend = False
End Function


Public Function CreateInboundBatchLog(BatchTypeKey As Long, SiteKey As Long, BatchNo As Long, _
 FormID As String, TotalRecords As Long, TotalErrors As Long) As Long
   Dim ssql As String
   Dim loRs As Recordset
   Dim loRs2 As Recordset
   Dim moRWork As mwReplicateWork
   On Error GoTo FunctionError
   ' New Batch Number
   If BatchNo < 1 Then
      CreateInboundBatchLog = -1
      Exit Function
   End If
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   With loRs
      .CursorLocation = adUseClient
      ssql = "select * from mwrBatchLogInbound where ID=-1"
      .Open ssql, goCon, adOpenDynamic, adLockOptimistic
      .AddNew
      .Fields("ID").value = goSession.MakePK("mwrBatchLogInbound")
      CreateInboundBatchLog = .Fields("ID").value
      .Fields("mwrBatchTypeKey").value = BatchTypeKey
      .Fields("mwcSitesKey").value = SiteKey
      .Fields("BatchNumber").value = BatchNo
      .Fields("DatagramFormID").value = FormID
      .Fields("DateReceived").value = Now()
      .Fields("TotalRecords").value = TotalRecords
      .Fields("TotalErrors").value = TotalErrors
      .Update
   End With
   '
   ' Create Modify Transaction to update AckBack on remote site
   '
   If mIncomingOutboundBatchKey > 0 Then
   
      goSession.RaiseWorkflowStatus "Creating AckBack transaction for Batch: " & CStr(mIncomingOutboundBatchKey) & ", from SiteKey: " & CStr(SiteKey)
   
      Set moRWork = New mwReplicateWork
      Set loRs2 = New Recordset
      With loRs2
         .Fields.Append "ID", adInteger, 4
         .Fields.Append "AckBackDate", loRs.Fields("DateReceived").Type, loRs.Fields("DateReceived").DefinedSize
         .Fields.Append "AckBackTotalErrors", adInteger, 4
         .Open
         .AddNew
         .Fields!ID = mIncomingOutboundBatchKey
         .Fields!AckBackTotalErrors = 0
         .Update
         .Fields!AckBackTotalErrors = TotalErrors
         .Fields!AckBackDate = Format(Now(), "YYYY-MM-DD hh:mm")
         If goSession.Site.SiteType = SITE_TYPE_SHIP Then
            moRWork.LogModifyChange mwrt_mwrBatchLogOutbound, .Fields, goSession.Site.SiteKey
         Else
            moRWork.LogModifyChange mwrt_mwrBatchLogOutbound, .Fields, SiteKey
         End If
         .Close
      End With
      Set loRs2 = Nothing
   End If
   CloseRecordset loRs
   goSession.RaiseWorkflowStatus "Processed Batch: " & CStr(BatchNo) & ", from SiteKey: " & CStr(SiteKey)
   Exit Function
FunctionError:
   If Err.Number = -2147217873 Then
      loRs.CancelUpdate
      CloseRecordset loRs
      CreateInboundBatchLog = 0
   Else
      goSession.RaiseError "General Error in mwSession.CreateInboundBatchLog. ", Err.Number, Err.Description
      CreateInboundBatchLog = -1
   End If
End Function


Private Function FetchBatchKeyTemplate(BatchTypeKey As Long) As String
   Dim loRs As Recordset
   Dim ssql As String
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   ssql = "select TemplateID from mwrBatchType where ID=" & BatchTypeKey
   loRs.Open ssql, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      FetchBatchKeyTemplate = ""
   Else
      If IsNull(loRs!TemplateID) Then
         FetchBatchKeyTemplate = ""
      Else
         FetchBatchKeyTemplate = loRs!TemplateID
      End If
   End If
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.FetchBatchKeyTemplate. ", Err.Number, Err.Description
   FetchBatchKeyTemplate = Null
End Function


Public Function CreateFleetBatchDatagrams(BatchTypeKey As Long) As Boolean
   MsgBox "function retired"
End Function

Private Function MarkFleetBatchDatagrams(BatchTypeKey As Long) As Long
   Dim ssql As String
   Dim loRsSites As Recordset
   Dim loRsChanges As Recordset
   Dim loRsColumns As Recordset
   Dim loRsDetail As Recordset
   Dim loRec As ConfigKeys
   Dim loform As mwForm
   Dim loDbWork As mwDataWork
   Dim strTemplate As String
   Dim IsError As Boolean
   Dim TotalRecords As Long
   Dim nBatchKey As Long
   Dim nBatchNo As Long
   On Error GoTo FunctionError
   
   MarkFleetBatchDatagrams = -1
   
   '
   ' Disable for demo loopback
   '
   'If goSession.User.DefaultTransport = mwt_DemoLoopback Then
   '   goSession.RaisePublicError "Error in mwReplicateBatchWork.MarkFleetBatchDatagrams, Fleet Transmission not available in Demo Loopback mode. "
   '   Exit Function
   'End If
   '
   ' Fetch Regional Settings
   '
   Set loDbWork = New mwDataWork
   
   '
   '  Create Datagram of fleet transactions
   '
   goSession.RaiseWorkflowStatus "creating Fleet Transmission for: " & BatchTypeKey
   '
   ' Fetch fleet changes...
   '
   Set loRsChanges = FetchChangesRS(BatchTypeKey, 0)
   If loRsChanges Is Nothing Then
      goSession.RaiseError "Error in mwReplicateBatchWork.MarkFleetBatchDatagrams, unable to open changes table. Operation Cancelled."
      CloseRecordset loRsChanges
      MarkFleetBatchDatagrams = -1
      Exit Function
   End If
   '
   '
   ' Any fleet changes ?
   '
   If loRsChanges.RecordCount < 1 Then
      mIsNoTransactions = True
      CloseRecordset loRsChanges
      goSession.RaiseWorkflowStatus "No transactions found for: " & BatchTypeKey
      MarkFleetBatchDatagrams = 0
      Exit Function
   End If
   '
   ' Get the Batch Key...
   '
   nBatchKey = CreateOutboundBatchLog(BatchTypeKey, nBatchNo)
   If nBatchKey < 1 Then
      CloseRecordset loRsChanges
      CloseRecordset loRsSites
      MarkFleetBatchDatagrams = -1
      Exit Function
   End If
   '
   ' Loop through change records...
   '
   TotalRecords = 0
   Do While Not loRsChanges.EOF
      '
      Select Case loRsChanges.Fields(RS_CL_ChangeType).value
         Case "A", "D", "M", "S", "L"
            '
            ' Update BatchKey in Change Log record
            '
            loRsChanges.Fields(RS_CL_mwrBatchLogOutboundKey).value = nBatchKey
            loRsChanges.Update
            loRsChanges.MoveNext
            TotalRecords = TotalRecords + 1
         Case Else
            MarkFleetBatchDatagrams = -1
            goSession.RaiseWorkflowStatus "Error in mwReplicateBatchWork. invalid change type code. Key: " & loRsChanges.Fields(RS_CL_ID).value
            loRsChanges.MoveNext
      End Select
   Loop
   CloseRecordset loRsChanges
   
   If Not FinalizeOutboundBatchlog(nBatchKey, "", TotalRecords) Then
      goSession.RaiseError "Error in mwReplicateBatchWork.MarkFleetBatchDatagrams, datagram sent - finalize log entry failed !"
      goCon.RollbackTrans
   End If
   mTotalFleetRecords = TotalRecords
   CloseRecordset loRsSites
   goSession.RaiseWorkflowStatus "Successfully completed Create Batch Datagrams."
   MarkFleetBatchDatagrams = nBatchKey
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwReplicateBatchWork.MarkFleetBatchDatagrams. ", Err.Number, Err.Description
   CloseRecordset loRsSites
   CloseRecordset loRsChanges
   MarkFleetBatchDatagrams = -1

End Function


Public Function ValidateIncomingSequence(SiteKey As Long, BatchTypeKey As Long, _
                ReceivedBatchNumber As Long, ByRef ExpectedBatchNumber As Long) As Boolean
   Dim loRs As Recordset
   Dim nBatchType As Long
   Dim nBatchNumber As Long
   Dim sTargetSiteID As String
   Dim ssql As String
   Dim s As String
   Dim b As Boolean
   On Error GoTo FunctionError
   '
   'Valid Replication Site
   '
   
   b = CBool(goSession.Site.GetShipProperty("", "IsReplicateSite", SiteKey))
   If Not b Then
      goSession.RaiseWorkflowStatus "Transmission recieved for Site where Replication is disabled. "
      ValidateIncomingSequence = False
      Exit Function
   End If
   '
   ' Received before ?
   '
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   ssql = "select * from mwrBatchLogInbound where mwcSitesKey=" & SiteKey & _
          " and mwrBatchTypeKey=" & BatchTypeKey & " order by BatchNumber"
   loRs.Open ssql, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      If ReceivedBatchNumber = 1 Then
         ValidateIncomingSequence = True
      Else
         ' File is empty, Batch 1 expected...
         goSession.RaiseWorkflowStatus "Transmission Received out of sequence: Received/Expected: " & _
           ReceivedBatchNumber & "/1"
         ValidateIncomingSequence = False
         gIsLeaveTransFile = True
         ExpectedBatchNumber = 1
         '
         ' Send Request Bit back...
         '
         RequestResend SiteKey, ExpectedBatchNumber, BatchTypeKey
      End If
   Else
      loRs.MoveLast
      If IsNull(loRs!BatchNumber) Then
         goSession.RaiseWorkflowStatus "Error in mwReplicateBatchWork.ValidateIncomingSequence, " & _
           " Last record has BatchNumber that is Null !"
         ValidateIncomingSequence = False
         ExpectedBatchNumber = -1
      ElseIf loRs!BatchNumber >= ReceivedBatchNumber Then
         goSession.RaiseWorkflowStatus "Transmission has already been received: " & ReceivedBatchNumber
         ExpectedBatchNumber = loRs!BatchNumber + 1
         ValidateIncomingSequence = False
      ElseIf loRs!BatchNumber <> ReceivedBatchNumber - 1 Then
         goSession.RaiseWorkflowStatus "Transmission received out of sequence. Received/Expected: " & _
            ReceivedBatchNumber & " / " & CStr(loRs!BatchNumber + 1)
         gIsLeaveTransFile = True
         ValidateIncomingSequence = False
         '
         ' Send Request Bit back...
         '
         RequestResend SiteKey, loRs!BatchNumber + 1, BatchTypeKey
         
      Else
         ValidateIncomingSequence = True
         ExpectedBatchNumber = ReceivedBatchNumber
      End If
   End If
   CloseRecordset loRs
   Exit Function
FunctionError:
   'Resume Next
   goSession.RaiseError "General Error in mwReplicateBatchWork.ValidateIncomingSequence. ", Err.Number, Err.Description
   ValidateIncomingSequence = False
   ExpectedBatchNumber = -1
End Function



Public Function UpdateAckBack(SiteKey As Long, BatchTypeKey As Long, BatchNo As Long) As Boolean
   Dim ssql As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   ' New Batch Number
   ssql = "select * from mwrBatchLogOutbound where mwcSitesKey=" & SiteKey & " and mwrBatchTypeKey=" & _
          BatchTypeKey & " and BatchNumber=" & BatchNo
   Set loRs = New Recordset
   loRs.Open ssql, goCon, adOpenDynamic, adLockOptimistic
   If loRs.RecordCount > 0 Then
      If IsNull(loRs!AckBackDate) Then
         loRs!AckBackDate = Now()
         loRs.Update
         UpdateAckBack = True
      Else
         UpdateAckBack = False
      End If
   Else
      UpdateAckBack = False
   End If
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwReplicateBatchWork.UpdateAckBack. ", Err.Number, Err.Description
   UpdateAckBack = -1
End Function


Public Function InstallFCB(FullFilename As String) As Boolean
   Dim loRsIF As Recordset
   Dim nSourceSiteKey As Long
   Dim loRsFileCab As Recordset
   Dim nExpBatchNo As Long
   Dim fso As FileSystemObject
   Dim ssql As String
   On Error Resume Next
   Set loRsIF = New Recordset
   loRsIF.Open FullFilename
   If Err Then
      goSession.RaiseWorkflowStatus "Error in mwReplicateBatchWork.InstallFCB. ", Err.Number, Err.Description
      goSession.RaisePublicError "Error in mwReplicateBatchWork.InstallFCB, Damaged File Received: " & FullFilename
      goSession.Workflow.ArchiveWorkflowItem FullFilename, ArchiveError, DELETE_AFTER_ARCHIVE
      InstallFCB = False
      Exit Function
   End If
   On Error GoTo FunctionError
   If loRsIF.RecordCount < 1 Then
      goSession.RaisePublicError "Error in mwReplicateBatchWork.InstallFCB, No records in file: " & FullFilename
      CloseRecordset loRsIF
      goSession.Workflow.ArchiveWorkflowItem FullFilename, ArchiveError, DELETE_AFTER_ARCHIVE
      InstallFCB = False
      Exit Function
   End If
   If IsNull(loRsIF!mwcSitesKey) Or IsNull(loRsIF!BatchType) Or IsNull(loRsIF!BatchNumber) Then
      goSession.RaisePublicError "Error in mwReplicateBatchWork.InstallFCB, Improper Null value in file: " & FullFilename
      CloseRecordset loRsIF
      goSession.Workflow.ArchiveWorkflowItem FullFilename, ArchiveError, DELETE_AFTER_ARCHIVE
      InstallFCB = False
      Exit Function
   End If
   
   
   If goSession.Site.SiteType = SITE_TYPE_SHIP Then
      nSourceSiteKey = goSession.Site.GetSiteKey(goSession.Site.TargetReplicateSiteID)
   Else
      nSourceSiteKey = loRsIF!mwcSitesKey
   End If
   
   
   '
   ' Valid Transmission ?
   '
   If Not ValidateIncomingSequence(nSourceSiteKey, loRsIF!BatchType, loRsIF!BatchNumber, nExpBatchNo) Then
      '
      ' Maybe request resend ?
      '
      
      '
      'goSession.RaisePublicError "Error in mwReplicateBatchWork.InstallFCB, Improper Null value in file: " & FullFilename
      If nExpBatchNo < loRsIF!BatchNumber Then
         ' Save for processing later
         'If Not goSession.Workflow.ArchiveWorkflowItem(FullFilename, ArchiveQueue, DELETE_AFTER_ARCHIVE) Then
         '   goSession.Workflow.RenameFileNowExtension (FullFilename)
         'End If
         InstallFCB = False
         CloseRecordset loRsIF
         Exit Function
      Else
         If nExpBatchNo = loRsIF!BatchNumber + 1 Then
            ' Already processed...but maybe again...
            goSession.RaiseWorkflowStatus "Reprocessing File Cabinet Update: " & goSession.Site.GetSiteName(loRsIF!mwcSitesKey) & " / " & loRsIF!BatchNumber
         Else
            ' Already processed...but maybe again...
            If Not goSession.Workflow.ArchiveWorkflowItem(FullFilename, ArchiveError, DELETE_AFTER_ARCHIVE) Then
               goSession.Workflow.RenameFileNowExtension (FullFilename)
            End If
            InstallFCB = True
            CloseRecordset loRsIF
            Exit Function
         End If
      End If
   End If
   '
   ' OK, install it...
   '
   Set loRsFileCab = New Recordset
   loRsFileCab.CursorLocation = adUseClient
   goSession.RaiseWorkflowStatus "Start Transaction: Install FCB from Site/Batch Number: " & goSession.Site.GetSiteName(loRsIF!mwcSitesKey) & " / " & loRsIF!BatchNumber
   goCon.BeginTrans
      Do While Not loRsIF.EOF
         'RS_FCB_ID = 0
         'RS_FCB_mwcSitesKey = 1
         'RS_FCB_BatchType = 2
         'RS_FCB_BatchKey = 3
         'RS_FCB_BatchNumber = 4
         'RS_FCB_DateMod = 5
         'RS_FCB_FileBlob = 6
         ssql = "select * from smFileCabinet where ID=" & loRsIF!ID
         loRsFileCab.Open ssql, goCon, adOpenDynamic, adLockOptimistic
      
         If loRsFileCab.EOF = True Then
            loRsFileCab.AddNew
            loRsFileCab!ID = loRsIF!ID
            loRsFileCab!mwcSitesKey = loRsIF!mwcSitesKey
            loRsFileCab!IsTempRecordInsert = True
            loRsFileCab.Update
            loRsFileCab.Close
            loRsFileCab.Open ssql, goCon, adOpenDynamic, adLockOptimistic

            goSession.RaiseWorkflowStatus "File Cabinet record " & CStr(loRsIF!ID) & " has not arrived yet. A Temporary record has been added to hold its place."
         End If
         
         If loRsFileCab.EOF = True Then
            goSession.RaiseWorkflowStatus "Could not complete the installation of the File Cabinet Package at this time, the new File Cabinet record " & CStr(loRsIF!ID) & " has not arrived yet. Package will remain in queue until it can be completed."
            CloseRecordset loRsFileCab
            CloseRecordset loRsIF
            goCon.RollbackTrans
            InstallFCB = False
            Exit Function
'         ElseIf loRsFileCab!DateMod > loRsIF!DateMod Then
'            goSession.RaiseWorkflowStatus "smFileCabinet record " & loRsIF!ID & " is newer than the incoming FCB record."
         ElseIf loRsFileCab!mwcSitesKey <> loRsIF!mwcSitesKey Then
            goSession.RaiseWorkflowStatus "smFileCabinet record " & loRsIF!ID & " has a different SiteKey than the incoming FCB record."
         Else
            '
            ' Update FileBlob record
            '
            loRsFileCab!FileBlob = loRsIF!FileBlob
            loRsFileCab!DateMod = loRsIF!DateMod
            loRsFileCab!mwrBatchLogOutboundKey = -1
            loRsFileCab.Update
            goSession.RaiseWorkflowStatus "smFileCabinet record " & loRsIF!ID & " has been updated."
         End If
         
         loRsFileCab.Close
         loRsIF.MoveNext
      Loop
      '
      ' Update Log
      '
      Set fso = New FileSystemObject
      loRsIF.MoveFirst
         
      mIncomingOutboundBatchKey = CLng(loRsIF!BatchKey)

      If CreateInboundBatchLog(loRsIF!BatchType, nSourceSiteKey, _
        loRsIF!BatchNumber, fso.GetBaseName(FullFilename), loRsIF.RecordCount, 0) < 0 Then
         goSession.RaisePublicError "Error in mwReplicateBatchWork.InstallFCB, Update inbound log failed: " & FullFilename
         CloseRecordset loRsIF
         goCon.RollbackTrans
         If Not goSession.Workflow.ArchiveWorkflowItem(FullFilename, ArchiveError, DELETE_AFTER_ARCHIVE) Then
            goSession.Workflow.RenameFileNowExtension (FullFilename)
         End If
         InstallFCB = False
         Exit Function
      End If
      
      
   goCon.CommitTrans
   '
   ' Cleanup
   '
   loRsIF.MoveFirst
   If goSession.Site.SiteType = SITE_TYPE_SHIP Then
      goSession.RaiseWorkflowStatus "Installed FCB Transmission from SHORE. Batch Number: " & loRsIF!BatchNumber
   Else
      goSession.RaiseWorkflowStatus "Installed FCB Transmission from " & goSession.Site.GetSiteName(loRsIF!mwcSitesKey) & _
       ", Batch Number: " & loRsIF!BatchNumber
   End If
   CloseRecordset loRsFileCab
   CloseRecordset loRsIF
   If Not goSession.Workflow.ArchiveWorkflowItem(FullFilename, ArchivedOK, DELETE_AFTER_ARCHIVE) Then
      goSession.Workflow.RenameFileNowExtension (FullFilename)
   End If
   InstallFCB = True
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwReplicateBatchWork.InstallFCB.", Err.Number, Err.Description
   If Not goSession.Workflow.ArchiveWorkflowItem(FullFilename, ArchiveError, DELETE_AFTER_ARCHIVE) Then
      goSession.Workflow.RenameFileNowExtension (FullFilename)
   End If
   CloseRecordset loRsFileCab
   CloseRecordset loRsIF
   InstallFCB = False
   On Error Resume Next
   goCon.RollbackTrans
End Function

Public Function InstallDAB(FullFilename As String) As Boolean
   Dim loRsIF As Recordset
   Dim nSourceSiteKey As Long
   Dim loRsDirAtt As Recordset
   Dim nExpBatchNo As Long
   Dim fso As FileSystemObject
   Dim ssql As String
   Dim nErrCount As Long
   Dim nBatchNumber As Long
   On Error Resume Next
   Set loRsIF = New Recordset
   loRsIF.Open FullFilename
   If Err Then
      goSession.RaiseWorkflowStatus "Error in mwReplicateBatchWork.InstallDAB. ", Err.Number, Err.Description
      goSession.RaisePublicError "Error in mwReplicateBatchWork.InstallDAB, Damaged File Received: " & FullFilename
      goSession.Workflow.ArchiveWorkflowItem FullFilename, ArchiveError, DELETE_AFTER_ARCHIVE
      InstallDAB = False
      Exit Function
   End If
   On Error GoTo FunctionError
   If loRsIF.RecordCount < 1 Then
      goSession.RaisePublicError "Error in msWorkstation.InstallDAB, No records in file: " & FullFilename
      CloseRecordset loRsIF
      goSession.Workflow.ArchiveWorkflowItem FullFilename, ArchiveError, DELETE_AFTER_ARCHIVE
      InstallDAB = False
      Exit Function
   End If
   If IsNull(loRsIF!mwcSitesKey) Or IsNull(loRsIF!BatchType) Or IsNull(loRsIF!BatchNumber) Then
      goSession.RaisePublicError "Error in mwReplicateBatchWork.InstallDAB, Improper Null value in file: " & FullFilename
      CloseRecordset loRsIF
      goSession.Workflow.ArchiveWorkflowItem FullFilename, ArchiveError, DELETE_AFTER_ARCHIVE
      InstallDAB = False
      Exit Function
   End If
   
   
   If goSession.Site.SiteType = SITE_TYPE_SHIP Then
      nSourceSiteKey = goSession.Site.GetSiteKey(goSession.Site.TargetReplicateSiteID)
   Else
      nSourceSiteKey = loRsIF!mwcSitesKey
   End If
   
   
   '
   ' Valid Transmission ?
   '
   If Not ValidateIncomingSequence(nSourceSiteKey, loRsIF!BatchType, loRsIF!BatchNumber, nExpBatchNo) Then
      '
      ' Maybe request resend ?
      '
      
      '
      'goSession.RaisePublicError "Error in mwReplicateBatchWork.InstallFCB, Improper Null value in file: " & FullFilename
      If nExpBatchNo < loRsIF!BatchNumber Then
         ' Save for processing later
'         If Not goSession.Workflow.ArchiveWorkflowItem(FullFilename, ArchiveQueue, DELETE_AFTER_ARCHIVE) Then
'            goSession.Workflow.RenameFileNowExtension (FullFilename)
'         End If
      Else
         ' Already processed...
         If Not goSession.Workflow.ArchiveWorkflowItem(FullFilename, ArchiveError, DELETE_AFTER_ARCHIVE) Then
            goSession.Workflow.RenameFileNowExtension (FullFilename)
         End If
      End If
      InstallDAB = False
      CloseRecordset loRsIF
      Exit Function
   End If
   '
   ' OK, install it...
   '
   Set loRsDirAtt = New Recordset
   loRsDirAtt.CursorLocation = adUseClient
   goSession.RaiseWorkflowStatus "Start Transaction: Install DAB from Site/Batch Number: " & goSession.Site.GetSiteName(loRsIF!mwcSitesKey) & " / " & loRsIF!BatchNumber
   goCon.BeginTrans
      Do While Not loRsIF.EOF
         'RS_FCB_ID = 0
         'RS_FCB_mwcSitesKey = 1
         'RS_FCB_BatchType = 2
         'RS_FCB_BatchKey = 3
         'RS_FCB_BatchNumber = 4
         'RS_FCB_FileBlob = 6
         ssql = "select * from smDirAttachment where ID=" & loRsIF!ID
         loRsDirAtt.Open ssql, goCon, adOpenDynamic, adLockOptimistic
      
         If loRsDirAtt.EOF = True Then
            goSession.RaiseWorkflowStatus "smDirAttachment record " & loRsIF!ID & " not found."
            nErrCount = nErrCount + 1
         Else
            '
            ' Update FileBlob record
            '
            loRsDirAtt!FileBlob = loRsIF!FileBlob
            
            If Not IsNull(loRsDirAtt!TargetLocation) Then
               InstallBLOBAttachment loRsDirAtt
            End If
            
            loRsDirAtt.Update
         End If
         
         loRsDirAtt.Close
         loRsIF.MoveNext
      Loop
      '
      ' Update Log
      '
      Set fso = New FileSystemObject
      loRsIF.MoveFirst
      nBatchNumber = loRsIF!BatchNumber
      If nErrCount > 0 Then
         nBatchNumber = 0
      End If
      mIncomingOutboundBatchKey = CLng(loRsIF!BatchKey)

      If CreateInboundBatchLog(loRsIF!BatchType, nSourceSiteKey, _
        nBatchNumber, fso.GetBaseName(FullFilename), loRsIF.RecordCount, 0) < 1 Then
         goSession.RaisePublicError "Error in mwReplicateBatchWork.InstallDAB, Update inbound log failed: " & FullFilename
         CloseRecordset loRsIF
         goCon.RollbackTrans
         If Not goSession.Workflow.ArchiveWorkflowItem(FullFilename, ArchiveError, DELETE_AFTER_ARCHIVE) Then
            goSession.Workflow.RenameFileNowExtension (FullFilename)
         End If
         InstallDAB = False
         Exit Function
      End If
   goCon.CommitTrans
   '
   ' Cleanup
   '
   loRsIF.MoveFirst
   If goSession.Site.SiteType = SITE_TYPE_SHIP Then
      goSession.RaiseWorkflowStatus "Installed DAB Transmission from SHORE. Batch Number: " & loRsIF!BatchNumber
   Else
      goSession.RaiseWorkflowStatus "Installed DAB Transmission from " & goSession.Site.GetSiteName(loRsIF!mwcSitesKey) & _
       ", Batch Number: " & loRsIF!BatchNumber
   End If
   CloseRecordset loRsDirAtt
   CloseRecordset loRsIF
   If Not goSession.Workflow.ArchiveWorkflowItem(FullFilename, ArchivedOK, DELETE_AFTER_ARCHIVE) Then
      goSession.Workflow.RenameFileNowExtension (FullFilename)
   End If
   InstallDAB = True
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwReplicateBatchWork.InstallDAB.", Err.Number, Err.Description
   If Not goSession.Workflow.ArchiveWorkflowItem(FullFilename, ArchiveError, DELETE_AFTER_ARCHIVE) Then
      goSession.Workflow.RenameFileNowExtension (FullFilename)
   End If
   CloseRecordset loRsDirAtt
   CloseRecordset loRsIF
   InstallDAB = False
   On Error Resume Next
   goCon.RollbackTrans
End Function

Private Function InstallBLOBAttachment(loRs As Recordset) As Boolean
   Dim loStream As ADODB.Stream
   Dim fso As FileSystemObject
   Dim TargetFName As String
   Dim SourceFile As String
   
   On Error GoTo FunctionError
   
   Set fso = New FileSystemObject
   Set loStream = New ADODB.Stream
   
   If Not IsNull(loRs!SourceFileLocation) And Not IsNull(loRs!TargetLocation) Then
      SourceFile = fso.GetFileName(loRs!SourceFileLocation)
      TargetFName = BuildPlaceFileName(loRs!TargetLocation, SourceFile)
   
      If fso.FileExists(TargetFName) Then
         fso.GetFile(TargetFName).Attributes = Normal
         fso.DeleteFile (TargetFName)
      End If
   
      loStream.Type = adTypeBinary
      loStream.Open
      
      loStream.Write loRs.Fields("FileBlob").value
      loStream.SaveToFile TargetFName, adSaveCreateOverWrite
      
      InstallBLOBAttachment = True
   Else
      InstallBLOBAttachment = False
   End If
   
   loRs!IsInstalled = -1
   
   KillObject fso
   KillObject loStream
   Exit Function
   
FunctionError:
   goSession.RaisePublicError "General Error in mwReplicateBatchWork.InstallBLOBAttachment", Err.Number, Err.Description
   KillObject fso
   KillObject loStream
   InstallBLOBAttachment = False

End Function
Private Function BuildPlaceFileName(PlaceID As String, FileName As String) As String
   Dim oRs As Recordset
   On Error GoTo FunctionError
   Set oRs = New Recordset
   oRs.CursorLocation = adUseClient
   oRs.Open "select * from mwcPlaces where PlaceID='" & PlaceID & "'", goCon, adOpenForwardOnly, adLockReadOnly
   If oRs.RecordCount < 1 Then
      goSession.RaisePublicError "General Error in mwReplicateBatchWork.BuildPlaceFileName: Could not find a record in mwcPlaces with PlaceID = " & PlaceID
   Else
      BuildPlaceFileName = ""
      If Not IsNull(oRs!DriveLetter) Then
         BuildPlaceFileName = BuildPlaceFileName & oRs!DriveLetter & ":\"
      End If
      If Not IsNull(oRs!RootPath) Then
         BuildPlaceFileName = BuildPlaceFileName & oRs!RootPath & "\"
      End If
      If Not IsNull(oRs!SubFolderPath) Then
         BuildPlaceFileName = BuildPlaceFileName & oRs!SubFolderPath & "\"
      End If
      BuildPlaceFileName = BuildPlaceFileName & FileName
   End If
   
   CloseRecordset oRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwReplicateBatchWork.BuildPlaceFileName. ", Err.Number, Err.Description
   BuildPlaceFileName = ""
   CloseRecordset oRs
End Function

Private Function IsSuspendTransmit(SiteID As String) As Boolean
   Dim loRs As Recordset
   Dim ssql As String
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   ssql = "select IsSuspendTransmit from mwcSites where SiteID='" & SiteID & "'"
   loRs.Open ssql, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      IsSuspendTransmit = False
   ElseIf IsNull(loRs!IsSuspendTransmit) Then
      IsSuspendTransmit = False
   ElseIf loRs!IsSuspendTransmit <> 0 Then
      IsSuspendTransmit = True
   Else
      IsSuspendTransmit = False
   End If
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwReplicateBatchWork.IsSuspendTransmit. ", Err.Number, Err.Description
End Function

Private Function RequestResend(SiteKey As Long, BatchNumber As Long, BatchTypeKey As Long) As Boolean
   Dim loform As mwForm
   Dim ssql As String
   Dim loRec As ConfigKeys
   Dim loRs As Recordset
   Dim sTarget As String
   Dim sEmail As String
   On Error GoTo FunctionError
   '
   ' Already sent it ?
   '
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   ssql = "select * from mwrBatchLogResendRequest where mwcSitesKey=" & SiteKey & _
     " and BatchNumber=" & BatchNumber & " and mwrBatchTypeKey = " & BatchTypeKey
   loRs.Open ssql, goCon, adOpenDynamic, adLockOptimistic
   If loRs.RecordCount < 1 Then
      '
      ' Set request resend bit on remote record...
      
      If goSession.Site.SiteType = SITE_TYPE_SHIP Then
'         sSQL = "update mwrBatchLogOutbound set AckBackResendRequest=1, AckBackRequestDate='" & _
'           Format(Now(), "YYYY-MM-DD hh:mm") & "'  where mwcSitesKey=" & goSession.Site.SiteKey & _
'           " and mwrBatchTypeKey=100 and BatchNumber=" & BatchNumber
         sTarget = goSession.Site.TargetReplicateSiteID
      Else
'         sSQL = "update mwrBatchLogOutbound set AckBackResendRequest=1, AckBackRequestDate='" & _
'           Format(Now(), "YYYY-MM-DD hh:mm") & "'  where mwcSitesKey=" & goSession.Site.SiteKey & _
'           " and mwrBatchTypeKey=100 and BatchNumber=" & BatchNumber
           sTarget = goSession.Site.GetSiteID(SiteKey)
   
      End If
      '
      Set loform = goSession.OpenNewForm()
      If Not loform.MakeNewForm("SYSTEM_CHANNEL", , , , sTarget) Then
         RequestResend = False
         Exit Function
      End If
      
      ssql = "update mwrBatchLogOutbound set AckBackResendRequest=1, AckBackRequestDate=#" & _
        Format(Now(), "yyyy-mm-dd hh:nn:ss") & "#  where mwcSitesKey=" & goSession.Site.SiteKey & _
        " and mwrBatchTypeKey=" & BatchTypeKey & " and BatchNumber=" & BatchNumber
      
      Set loRec = loform.DataForm.AddRecord
      With loRec
         .Add "Operation", "X"
         .Add "1", ssql
      End With
      '
      ' Write the message...
      '
      If Not loform.DataForm.WriteDataForm() Then
         RequestResend = False
         Exit Function
      End If
      '
      ' Transmit the message
      '
      loform.DeleteAfterSend = True
      
      If goSession.Site.SiteType = SITE_TYPE_SHIP Then
         If Not loform.Send(, mw_SendToShore, sTarget) Then
            ' Error sending email
            RequestResend = False
         Else
            RequestResend = True
         End If
      Else
         sEmail = goSession.Site.GetShipProperty("", "ShipWorkflowSendToAddress", SiteKey)
         If Not loform.Send(sEmail, mw_SendToShip, sTarget) Then
            ' Error sending email
            RequestResend = False
         Else
            RequestResend = True
         End If
      End If
      loform.TearDown
      Set loform = Nothing
      '
      '
      If RequestResend Then
         With loRs
            .AddNew
            .Fields!ID = goSession.MakePK("mwrBatchlogResendRequest")
            .Fields!mwcSitesKey = SiteKey
            .Fields!BatchNumber = BatchNumber
            .Fields!DateRequested = Now()
            .Fields!NoTimesSent = 1
            .Fields!mwrBatchTypeKey = BatchTypeKey
            .Update
         End With
      End If
      goSession.RaiseWorkflowStatus "Sent Resend Request to Site/Batch Number: " & CStr(SiteKey) & "/" & CStr(BatchNumber)
   Else
      RequestResend = True
   End If
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwReplicateBatchWork.RequestResend. ", Err.Number, Err.Description
End Function

Public Function RemakeTransmissionFile(BatchKey As Long) As Boolean
   Dim loRs As Recordset
   Dim loRsChanges As Recordset
   Dim loform As mwForm
   Dim loRec As ConfigKeys
   
   Dim ssql As String
   On Error GoTo FunctionError
   ssql = "select * from mwrBatchLogOutbound where ID=" & BatchKey
   Set loRs = New Recordset
   loRs.Open ssql, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      goSession.RaisePublicError "Error in mwReplicateBatchWork.RemakeTransmissionFile, Batch Note Found: " & BatchKey
      CloseRecordset loRs
      RemakeTransmissionFile = False
      Exit Function
   End If
   '
   mStartChangeLogKey = 0
   mEndChangeLogKey = 0
   
   Set loRsChanges = FetchAuditRS(loRs!ID, ZeroNull(loRs!mwrBatchLogOutboundFleetKey), _
                                  ZeroNull(loRs!StartChangeLogKey), ZeroNull(loRs!EndChangeLogKey))
   DoEvents
   
   If loRsChanges.RecordCount > 0 Then
   
      Set loform = goSession.OpenNewForm
      If Not loform.MakeNewForm("BRS", goSession.Workflow.ArchiveOkZipFile, , loRs!DatagramFormID) Then
         CloseRecordset loRsChanges
         CloseRecordset loRs
         RemakeTransmissionFile = False
         Exit Function
      End If
      '
      ' Populate Target Info in Header...
      '
      
      With loform.DataForm.Records(1).ConfigKeys
         .Add "RemakeTransmission", "True", "RemakeTransmission"
         If goSession.Site.SiteType <> SITE_TYPE_SHORE Then
            ' Ship to Shore
            .Add "TargetSiteID", goSession.Site.TargetReplicateSiteID, "TargetSiteID"
            .Add "TargetSiteKey", goSession.Site.GetSiteKey(goSession.Site.TargetReplicateSiteID), "TargetSiteKey"
         Else
            ' Shore to Ship
            .Add "TargetSiteID", goSession.Site.GetSiteID(loRs!mwcSitesKey), "TargetSiteID"
            .Add "TargetSiteKey", loRs!mwcSitesKey, "TargetSiteKey"
         End If
         .Add "BatchKey", loRs!ID, "BatchKey"
         .Add "BatchType", loRs!mwrBatchTypeKey, "BatchType"
         .Add "BatchNumber", loRs!BatchNumber, "BatchNumber"
         If Not IsNull(loRs!mwrBatchLogOutboundFleetKey) Then
            .Add "FleetBatchNumber", loRs!mwrBatchLogOutboundFleetKey, "FleetBatchNumber"
         Else
            .Add "FleetBatchNumber", 0, "FleetBatchNumber"
         End If
      End With
      '
      ' Loop through change records...
      '
      goSession.RaiseWorkflowStatus "...Rewriting datagram " & loRs!DatagramFormID & " with " & loRsChanges.RecordCount & " Changes to " & goSession.Site.GetSiteName(loRs!mwcSitesKey)
      Do While Not loRsChanges.EOF         '
         DoEvents
         Select Case loRsChanges.Fields(RS_CL_ChangeType).value
            Case "D"
               ' Delete Transaction
               With loform.DataForm.AddRecord
                  .Add "OPERATION", "D", "OPERATION"
                  .Add "TABLE", loRsChanges.Fields(RS_CT_Tablename).value, "TABLE"
                  .Add "ID", loRsChanges.Fields(RS_CL_RecordKey).value, "ID"
               End With
            Case "A"
               ' Add Transaction
               Set loRec = loform.DataForm.AddRecord
               loRec.Add "OPERATION", "A", "OPERATION"
               loRec.Add "TABLE", loRsChanges.Fields(RS_CT_Tablename).value, "TABLE"
               If Not WriteAuditChangeColumns(loRec, loRsChanges.Fields(RS_CL_ID).value) Then
                  CloseRecordset loRs
                  CloseRecordset loRsChanges
                  goSession.RaiseWorkflowStatus "Error Creating Batch Datagrams. Process cancelled."
                  goCon.RollbackTrans
                  Exit Function
               End If
            Case "M"
               ' Modify Transaction
               Set loRec = loform.DataForm.AddRecord
               loRec.Add "OPERATION", "M", "OPERATION"
               loRec.Add "TABLE", loRsChanges.Fields(RS_CT_Tablename).value, "TABLE"
               loRec.Add "ID", loRsChanges.Fields(RS_CL_RecordKey).value, "ID"
               If Not WriteAuditChangeColumns(loRec, loRsChanges.Fields(RS_CL_ID).value) Then
                  CloseRecordset loRs
                  CloseRecordset loRsChanges
                  goSession.RaiseWorkflowStatus "Error Creating Batch Datagrams. Process cancelled."
                  goCon.RollbackTrans
                  Exit Function
               End If
               
            Case "S"
               ' Schema Transaction
               Set loRec = loform.DataForm.AddRecord
               
               loRec.Add "OPERATION", "S", "OPERATION"
               loRec.Add "TABLE", loRsChanges.Fields(RS_CT_Tablename).value, "TABLE"
               If Not WriteAuditChangeColumns(loRec, loRsChanges.Fields(RS_CL_ID).value) Then
                  CloseRecordset loRs
                  CloseRecordset loRsChanges
                  goSession.RaiseWorkflowStatus "Error Creating Batch Datagrams. Process cancelled."
                  goCon.RollbackTrans
                  Exit Function
               End If
               
            Case "L"
               ' send Log Transaction email (like Schema Transaction)
               Set loRec = loform.DataForm.AddRecord
               
               loRec.Add "OPERATION", "L", "OPERATION"
               loRec.Add "TABLE", loRsChanges.Fields(RS_CT_Tablename).value, "TABLE"
               If Not WriteAuditChangeColumns(loRec, loRsChanges.Fields(RS_CL_ID).value) Then
                  CloseRecordset loRs
                  CloseRecordset loRsChanges
                  goSession.RaiseWorkflowStatus "Error Creating Batch Datagrams. Process cancelled."
                  goCon.RollbackTrans
                  Exit Function
               End If
               
            Case Else
               goSession.RaiseWorkflowStatus "Warning in mwReplicateBatchWork. invalid change type code. Key: " & loRsChanges.Fields(RS_CL_ID).value
         End Select
         '
         ' Update BatchKey in Change Log record if it isn't already set to the Fleet Batch Key
         '
         If IsNull(loRsChanges.Fields(RS_CL_mwrBatchLogOutboundKey).value) Then
            loRsChanges.Fields(RS_CL_mwrBatchLogOutboundKey).value = BatchKey
            loRsChanges.Update
         End If
         loRsChanges.MoveNext
      Loop
      '
      ' Create Integrity Trailor Record
      '
      With loform.DataForm.AddRecord
         .Add "Operation", "I"
         .Add "COMMENT", "Batch Integrity Record"
         .Add "SiteID", goSession.Site.SiteID, "SiteID"
         If goSession.Site.SiteType <> SITE_TYPE_SHORE Then
            ' Ship to shore
            .Add "TargetSiteID", goSession.Site.TargetReplicateSiteID, "TargetSiteID"
            .Add "TargetSiteKey", goSession.Site.GetSiteKey(goSession.Site.TargetReplicateSiteID), "TargetSiteKey"
         Else
            ' Shore to ship
            .Add "TargetSiteID", goSession.Site.GetSiteID(loRs!mwcSitesKey), "TargetSiteID"
            .Add "TargetSiteKey", loRs!mwcSitesKey, "TargetSiteKey"
         End If
         .Add "BatchType", loRs!mwrBatchTypeKey, "BatchType"
         .Add "BatchNumber", loRs!BatchNumber, "BatchNumber"
         ' added by TJM
         If Not IsNull(loRs!mwrBatchLogOutboundFleetKey) Then
            .Add "FleetBatchNumber", loRs!mwrBatchLogOutboundFleetKey, "FleetBatchNumber"
         Else
            .Add "FleetBatchNumber", 0, "FleetBatchNumber"
         End If
         .Add "BatchKey", CStr(loRs!ID), "BatchKey"
         .Add "TotalRecords", loform.DataForm.Records.Count
      End With
      
      '
      '
      DoEvents
      If Not loform.DataForm.WriteDataForm(goSession.Workflow.ArchiveOkZipFile) Then
         goSession.RaiseError "Error in mwReplicateBatchWork.RemakeTransmissionFile, WriteDataForm Failed. Operation Cancelled."
         CloseRecordset loRsChanges
         CloseRecordset loRs
         goCon.RollbackTrans
         loform.TearDown
         RemakeTransmissionFile = False
         Exit Function
      End If
   Else
      RemakeTransmissionFileFromArchive BatchKey
   End If
   CloseRecordset loRsChanges
   
   RemakeTransmissionFile = True
   Exit Function
FunctionError:
'Resume Next
   goSession.RaisePublicError "General Error in mwReplicateBatchWork.RemakeTransmissionFile. ", Err.Number, Err.Description
End Function

Public Function SendFileCabinetPackages() As Boolean
   'By N.Angelakis On 12th June 2009
   'DEV-1716 Ship Certificates - Fleet File Cabinets
   ' Moved to mwReplicateBatchWork and modified by T.McSally 12 July 2009
   
   Dim ssql As String
   Dim loRs As Recordset
   
   Const SEND_CDROM_FALSE = False
'Const BATCH_TYPE_SITE_DATAGRAM = 100
'Const BATCH_TYPE_FILECAB_MODEM = 101
'Const BATCH_TYPE_FILECAB_MEDIA = 102

   Dim nBatchKey As Long
   Dim nBatchType As Long
   Dim nBatchNo As Long
   Dim sFileName As String
   Dim sFullFilename As String
   Dim sEmailAddress As String
   Dim loRsFAB As Recordset
   Dim loRsBLOB As Recordset
   Dim fso As FileSystemObject
   Dim IsSentByMedia As Boolean
   Dim nTransportMethod As Long
   Dim sTargetSiteID As String
   
   On Error GoTo FunctionError
   
   goSession.RaiseWorkflowStatus "Send File Cabinet Packages: Creating outbound File Cabinet Packages"
   
   IsSentByMedia = False
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   
   ssql = "SELECT  mwcSites.SiteName, mwcSites.siteID, mwcSites.PackageSizeThreshold, " & _
         "         smFileCabinet.FileBLOB, smFileCabinet.ID, smFileCabinet.mwcSitesKey, smFileCabinet.IsSendNow, smFileCabinet.mwrBatchLogOutboundKey  " & _
         " FROM    mwcSites, smFileCabinet " & _
         " WHERE   smFileCabinet.mwcSitesKey=mwcSites.ID" & _
         " AND     smFileCabinet.mwrBatchLogOutboundKey Is Null " & _
         " AND     smFileCabinet.FullFileName Is Not Null  " & _
         " AND     smFileCabinet.FileBlob Is Not Null" & _
         " AND     smFileCabinet.mwcSitesKeyOrig=" & goSession.Site.SiteKey & _
         " AND     (mwcSites.IsSuspendTransmit = 0 OR mwcSites.IsSuspendTransmit Is Null) " & _
         " AND     mwcSites.PackageSizeThreshold Is Not Null"
 
   loRs.Open ssql, goCon, adOpenDynamic, adLockOptimistic
   
   While Not loRs.EOF
      If ZeroNull(loRs("FileBlob").ActualSize) <= ZeroNull(loRs!PackageSizeThreshold) Then
         
         If goSession.Site.GetShipProperty(loRs!mwcSitesKey, "SendAttachmentsByMedia") = "-1" Or _
            goSession.Site.GetShipProperty(loRs!mwcSitesKey, "SendAttachmentsByMedia") = "1" Then
            
            nBatchType = BATCH_TYPE_FILECAB_MEDIA
            nTransportMethod = mwt_Media
         Else
            nTransportMethod = GetTransportMethod(loRs!mwcSitesKey)
            
            If nTransportMethod = mwt_Media Then
               nBatchType = BATCH_TYPE_FILECAB_MEDIA
            Else
               nBatchType = BATCH_TYPE_FILECAB_MODEM
            End If
         End If
  
'         If goSession.Site.SiteKey = loRs!mwcSitesKey Then
'            nTransportMethod = mwt_DemoLoopback
'         End If
         
         Set loRsFAB = New Recordset
         With loRsFAB.Fields
            .Append "ID", adInteger, 4
            .Append "mwcSitesKey", adInteger, 4
            .Append "BatchType", adInteger, 4
            .Append "BatchKey", adInteger, 4
            .Append "BatchNumber", adInteger, 4
            .Append "DateMod", adDate
            .Append "FileBlob", adLongVarBinary, LONG_BIN_DEFINED_SIZE, adFldLong
         End With
         loRsFAB.Open

         goCon.BeginTrans
            If goSession.Site.SiteType = SITE_TYPE_SHORE Then
               nBatchKey = CreateOutboundBatchLog(nBatchType, nBatchNo, loRs!mwcSitesKey)
            Else
               nBatchKey = CreateOutboundBatchLog(nBatchType, nBatchNo, _
                  goSession.Site.GetSiteKey(goSession.Site.TargetReplicateSiteID))
            End If
            If nBatchKey < 1 Then
               goCon.RollbackTrans
               goSession.RaiseWorkflowStatus "Cancelling Automatic File Cabinet Replication due to an error in CreateOutboundBatchLog."
               CloseRecordset loRsFAB
               CloseRecordset loRs
               SendFileCabinetPackages = False
               Exit Function
            End If
            
            loRsFAB.AddNew
            loRsFAB!ID = loRs!ID
            loRsFAB!mwcSitesKey = loRs!mwcSitesKey
            loRsFAB!BatchType = nBatchType
            loRsFAB!BatchKey = nBatchKey
            loRsFAB!BatchNumber = nBatchNo
            loRsFAB!DateMod = Now 'moRsGrid!DateMod

            Set loRsBLOB = FetchBLOBRs(loRs!ID)
            If loRsBLOB.RecordCount = 1 Then
               loRsFAB!FileBlob = loRs!FileBlob
            End If
            CloseRecordset loRsBLOB
               
            ' Save/Archive File
            ' Update Outboundlog
            ' Commit Transaction
            loRsFAB.Update
            
            loRs!mwrBatchLogOutboundKey = nBatchKey
            loRs!IsSendNow = 1
            loRs.Update
         

            If goSession.Site.SiteType = SITE_TYPE_SHIP Then
               sFileName = goSession.Site.ShortSiteID & _
                 goSession.Site.GetShipProperty(goSession.Site.TargetReplicateSiteID, "ShortSiteID") & _
                 Format(nBatchNo, "0000")
            Else
               sFileName = goSession.Site.ShortSiteID & _
                 goSession.Site.GetShipProperty("", "ShortSiteID", loRsFAB!mwcSitesKey) & _
                 Format(nBatchNo, "0000")
            End If
            If nBatchType = BATCH_TYPE_FILECAB_MODEM Then
               ' modem
               sFullFilename = goSession.Workflow.ArchiveOkZipFile & "\" & sFileName & ".FMB"
            Else
               sFullFilename = goSession.Workflow.ArchiveOkZipFile & "\" & sFileName & ".FCB"
            End If
            Set fso = New FileSystemObject
            If fso.FileExists(sFullFilename) Then
               goSession.Workflow.RenameFileNowExtension sFullFilename
            End If
            loRsFAB.Save sFullFilename, adPersistADTG
            If Not FinalizeOutboundBatchlog(nBatchKey, sFileName, loRsFAB.RecordCount) Then
               goCon.RollbackTrans
               CloseRecordset loRsFAB
               CloseRecordset loRs
               SendFileCabinetPackages = False
               Exit Function
            End If
         goCon.CommitTrans


         '
         ' Try and Send it...
         '
         
         If goSession.Site.SiteType = SITE_TYPE_SHIP Then
            sTargetSiteID = goSession.Site.TargetReplicateSiteID
            sEmailAddress = goSession.Site.GetShipProperty(sTargetSiteID, "WorkflowSendToAddress")
         Else
            sTargetSiteID = loRs!SiteID
            sEmailAddress = goSession.Site.GetShipProperty(sTargetSiteID, "ShipWorkflowSendToAddress")
         End If
            
         
         Select Case nTransportMethod
            Case mwt_ElectronicMail
            
               If Not goSession.Workflow.SendByEmail(sFullFilename, "File Cabinet Transmission " & str(nBatchNo) & " From " & goSession.Site.SiteName, sEmailAddress, , loRs!mwcSitesKey) Then
                  goSession.RaiseWorkflowStatus "Error in mwWorkstation.FrmWorkflowAgent.SendFileCabinetPackages. File was prepared but not sent for: " & loRs!SiteName
               Else
                  goSession.RaiseWorkflowStatus "File Cabinet Transmission " & str(nBatchNo) & " sent to : " & loRs!SiteName
               End If
            
            Case mwt_Media, mwt_TransportContainer
            
'Public Function SendByMedia(FullFilename As String, Optional Zipit As Boolean, _
'  Optional ExplicitPath As String, Optional ChooseTargetPath As Boolean, _
'  Optional TargetSiteID As String, _
'  Optional IsSendByCDROM As Boolean, Optional TargetSiteKey As Long) As Boolean
            
               If Not goSession.Workflow.SendByMedia(sFullFilename, ZIP_IT, , , sTargetSiteID, SEND_CDROM_FALSE, loRs!mwcSitesKey) Then
                  goSession.RaiseWorkflowStatus "Error in mwReplicateBatchWork.SendFileCabinetPackages. File was prepared but not sent for: " & loRs!SiteName
               Else
                  goSession.RaiseWorkflowStatus "File Cabinet Transmission " & str(nBatchNo) & " successfully sent to : " & loRs!SiteName
               End If
            Case mwt_DemoLoopback
               goSession.Workflow.SendByDemoLoopback sFullFilename
            
         End Select
         
'         If nBatchType = BATCH_TYPE_FILECAB_MEDIA Then
'
'            If goSession.User.DefaultTransport = mwt_Media Then
'               '
'               '
'               '
'               If Not goSession.Workflow.SendByMedia(sFullFilename, ZIP_IT, , , loRs!mwcSitesKey, SEND_CDROM_FALSE, loRs!mwcSitesKey) Then
'                  goSession.RaiseWorkflowStatus "Error in mwReplicateBatchWork.SendFileCabinetPackages. File was prepared but not sent for: " & loRs!SiteName
'               Else
'                  goSession.RaiseWorkflowStatus "File Cabinet Transmission " & str(nBatchNo) & " successfully sent to : " & loRs!SiteName
'               End If
'            Else
'               If goSession.Site.SiteType = SITE_TYPE_SHIP Then
'                  sEmailAddress = goSession.Site.GetShipProperty("", "WorkflowSendToAddress", loRs!mwcSitesKey)
'               Else
'                  sEmailAddress = goSession.Site.GetShipProperty("", "ShipWorkflowSendToAddress", loRs!mwcSitesKey)
'               End If
'
'
'               If Not goSession.Workflow.SendByEmail(sFullFilename, "File Cabinet Transmission " & str(nBatchNo) & " From " & goSession.Site.SiteName, sEmailAddress, , loRs!mwcSitesKey) Then
'                  goSession.RaiseWorkflowStatus "Error in mwWorkstation.FrmWorkflowAgent.SendFileCabinetPackages. File was prepared but not sent for: " & loRs!SiteName
'               Else
'                  goSession.RaiseWorkflowStatus "File Cabinet Transmission successfully " & str(nBatchNo) & " sent to : " & loRs!SiteName
'               End If
'            End If
'         Else
'            If Not goSession.Workflow.SendByMedia(sFullFilename, ZIP_IT, , , loRs!SiteID, CDROM_IT, loRs!mwcSitesKey) Then
'               goSession.RaiseWorkflowStatus "Error in mwWorkstation.FrmWorkflowAgent.SendFileCabinetPackages. File was prepared but not sent for: " & loRs!SiteName
'            Else
'               goSession.RaiseWorkflowStatus "File Cabinet Transmission " & str(nBatchNo) & " successfully sent to : " & loRs!SiteName
'               IsSentByMedia = True
'            End If
'         End If
         
         
         '
         ' cleanup...
         '
         CloseRecordset loRsFAB
      End If
      loRs.MoveNext
   Wend
   
   If IsSentByMedia = True Then
      goSession.RaiseWorkflowStatus "File Cabinet Transmission Files for CDROM Transmission may be found in Site-specific folders under: " & goSession.Workflow.SendByCdromFolder
   End If
    
   '
   ' cleanup...
   '
   CloseRecordset loRs
   SendFileCabinetPackages = True
   
   DoEvents
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwReplicateBatchWork.SendFileCabinetPackages. ", Err.Number, Err.Description
   CloseRecordset loRsBLOB
   CloseRecordset loRsFAB
   CloseRecordset loRs
   SendFileCabinetPackages = False
End Function

Private Function GetTransportMethod(TargetSiteKey As Long) As Long
   Dim nTransport As Long
   
   On Error GoTo FunctionError
   
'Public Enum mwt_TransportMethod
'   mwt_TransportContainer = 0
'   mwt_ElectronicMail = 1
'   mwt_Media = 2
'   mwt_DemoLoopback = 3
'End Enum
   
   If IsNumeric(goSession.Site.GetShipProperty("", "DefaultTransport", TargetSiteKey)) Then
      GetTransportMethod = CLng(goSession.Site.GetShipProperty("", "DefaultTransport", TargetSiteKey))
   End If
   
   If GetTransportMethod < 1 Then
      GetTransportMethod = goSession.User.DefaultTransport
   End If
   
   Exit Function
FunctionError:
   GetTransportMethod = 0
End Function

Private Function FetchBLOBRs(smFileCabKey As Long) As Recordset
   'By N.Angelakis On 12th June 2009
   'DEV-1716 Ship Certificates - Fleet File Cabinets
   Dim ssql As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   
   ssql = "SELECT FileBLOB FROM smFileCabinet WHERE ID= " & smFileCabKey
   
   loRs.Open ssql, goCon, adOpenDynamic, adLockOptimistic
   
   Set FetchBLOBRs = loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwReplicateBatchWork.FetchBLOBRs: ", Err.Number, Err.Description
   CloseRecordset loRs
End Function

Public Function CreateLogReplicationRequest(SendToUserKey As Long, SendToUserEmail As String, _
      ReqFromUserKey As String, ReqFromUserEmail As String) As Boolean

'Public Function CreateLogReplicationRequest(filepath As String, SchemaName As String) As Boolean
   Dim loStream As ADODB.Stream
   Dim fso As FileSystemObject
   Dim loRs As Recordset
   Dim sStreamContents As String
   On Error GoTo FunctionError
   '
   ' create datagram to request log (from site)
   '
   'Set fso = New FileSystemObject
   'If Not fso.FileExists(filepath) Then      ' get the long log file name
   '   goSession.RaisePublicError "Error in mwSession.mwSchemaUpdateWork.CreateLogReplicationRequest, File Not Found: " & filepath
   '   KillObject fso
   '   CreateLogReplicationRequest = False
   '   Exit Function
   'End If

   'Set loStream = New ADODB.Stream
   '
   'loStream.Type = adTypeText
   'loStream.Charset = "ascii"
   'loStream.Open
   'loStream.LoadFromFile filepath
   'sStreamContents = loStream.ReadText

   'If loStream.Size > 30000 Then
   '   goSession.RaisePublicError "Schema " & SchemaName & " is too large to replicate. Maximum size is 30,000 bytes."
   '   KillObject fso
   '   KillObject loStream
   '   CreateLogReplicationRequest = False
   '   Exit Function
   'End If

   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Fields.Append "SchemaName", adVarWChar, 20
   loRs.Fields.Append "SendToUserKey", adInteger
   loRs.Fields.Append "SendToUserEmail", adVarWChar, 200
'   loRs.Fields.Append "SchemaContents", adLongVarWChar, 50000
'   loRs.Fields.Append "SchemaContents", adLongVarWChar, (loStream.Size + 10) * 2
   'loRs.Fields.Append "SchemaContents", adLongVarWChar, loStream.Size
   loRs.Open
      
   loRs.AddNew
   'loRs!SchemaName = SchemaName
   loRs!SchemaName = "Log Request"
   loRs!SendToUserKey = SendToUserKey
   loRs!SendToUserEmail = SendToUserEmail
   'loRs!SchemaContents = sStreamContents
   loRs.Update
   
   CreateLogReplicationRequest = goSession.ReplicateWork.LogLogRequest(loRs.Fields)
   
   CloseRecordset loRs

   KillObject fso
   KillObject loStream

   Exit Function

FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwSchemaUpdateWork.AddRecord. ", Err.Number, Err.Description
   CloseRecordset loRs
   KillObject fso
   KillObject loStream
   CreateLogReplicationRequest = False
End Function


Public Function RemakeTransmissionFileFromArchive(BatchKey As Long) As Boolean
   Dim loRs As Recordset
   Dim loRsChanges As Recordset
   Dim loform As mwForm
   Dim loRec As ConfigKeys
   
   Dim nMinChangeLogKey As Long
   Dim nMaxChangeLogKey As Long
   Dim ssql As String
   Dim sArchiveTableName As String
   Dim sArchiveTableChildName As String
   Dim sAltDBConnectionString As String
   Dim loArchiveDetail As Recordset
   Dim lorsChangeTable As Recordset
   Dim dbCon As Connection
   Dim sTableName As String
   
   On Error GoTo FunctionError
   ssql = "select * from mwrBatchLogOutbound where ID=" & BatchKey
   Set loRs = New Recordset
   loRs.Open ssql, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      goSession.RaisePublicError "Error in mwReplicateBatchWork.RemakeTransmissionFileFromArchive, Batch Note Found: " & BatchKey
      CloseRecordset loRs
      RemakeTransmissionFileFromArchive = False
      Exit Function
   End If
   '
   mStartChangeLogKey = 0
   mEndChangeLogKey = 0
   
   Set loArchiveDetail = New Recordset
   loArchiveDetail.CursorLocation = adUseClient
   
   ssql = "SELECT * FROM snArchiveLogDetail where mwrBatchLogOutboundKey = " & BatchKey
   loArchiveDetail.Open ssql, goCon, adOpenStatic, adLockReadOnly
   
   If Not IsRecordLoaded(loArchiveDetail) Then
      Exit Function
   Else
      sArchiveTableName = UCase(loArchiveDetail!ArchiveTableName)
      sArchiveTableChildName = Replace(sArchiveTableName, "MWRAUDITLOG_", "MWRAUDITLOGCOLUMN_")
      If Not IsNull(loArchiveDetail!DBConnectString) Then
         sAltDBConnectionString = loArchiveDetail!DBConnectString
      End If
      If sAltDBConnectionString = "" Then
         Set dbCon = goSession.DBConnection
      Else
         Set dbCon = New ADODB.Connection
         dbCon.Open sAltDBConnectionString
      End If
   End If
   Set loArchiveDetail = Nothing
   
   Set lorsChangeTable = New Recordset
   lorsChangeTable.CursorLocation = adUseClient
   
   ssql = "SELECT * FROM mwrChangeTable"
   lorsChangeTable.Open ssql, goCon, adOpenStatic, adLockReadOnly
   '
   
   ssql = "SELECT * " & _
      " FROM " & sArchiveTableName & " " & _
      " WHERE (mwrBatchLogOutboundKey=" & _
      BatchKey & " or mwrBatchLogOutboundKey=" & ZeroNull(loRs!mwrBatchLogOutboundFleetKey) & ") "
      
   nMinChangeLogKey = ZeroNull(loRs!StartChangeLogKey)
   If nMinChangeLogKey > 0 Then
      ssql = ssql & " AND ID >= " & nMinChangeLogKey
   End If
   
   nMaxChangeLogKey = ZeroNull(loRs!EndChangeLogKey)
   If nMaxChangeLogKey > 0 Then
      ssql = ssql & " AND ID <= " & nMaxChangeLogKey
   End If
   
   ssql = ssql & " ORDER BY ID"
   
   Set loRsChanges = New Recordset
   loRsChanges.CursorLocation = adUseClient
   loRsChanges.Open ssql, dbCon, adOpenDynamic, adLockOptimistic
   If Not (loRsChanges.EOF Or loRsChanges.BOF) Then
      loRsChanges.MoveFirst
   End If
   
   DoEvents
   
   If loRsChanges.RecordCount > 0 Then
   
      Set loform = goSession.OpenNewForm
      If Not loform.MakeNewForm("BRS", goSession.Workflow.ArchiveOkZipFile, , loRs!DatagramFormID) Then
         CloseRecordset loRsChanges
         CloseRecordset loRs
         CloseRecordset lorsChangeTable
         RemakeTransmissionFileFromArchive = False
         Exit Function
      End If
      '
      ' Populate Target Info in Header...
      '
      
      With loform.DataForm.Records(1).ConfigKeys
         .Add "RemakeTransmission", "True", "RemakeTransmission"
         If goSession.Site.SiteType <> SITE_TYPE_SHORE Then
            ' Ship to Shore
            .Add "TargetSiteID", goSession.Site.TargetReplicateSiteID, "TargetSiteID"
            .Add "TargetSiteKey", goSession.Site.GetSiteKey(goSession.Site.TargetReplicateSiteID), "TargetSiteKey"
         Else
            ' Shore to Ship
            .Add "TargetSiteID", goSession.Site.GetSiteID(loRs!mwcSitesKey), "TargetSiteID"
            .Add "TargetSiteKey", loRs!mwcSitesKey, "TargetSiteKey"
         End If
         .Add "BatchKey", loRs!ID, "BatchKey"
         .Add "BatchType", loRs!mwrBatchTypeKey, "BatchType"
         .Add "BatchNumber", loRs!BatchNumber, "BatchNumber"
         If Not IsNull(loRs!mwrBatchLogOutboundFleetKey) Then
            .Add "FleetBatchNumber", loRs!mwrBatchLogOutboundFleetKey, "FleetBatchNumber"
         Else
            .Add "FleetBatchNumber", 0, "FleetBatchNumber"
         End If
      End With
      '
      ' Loop through change records...
      '
      goSession.RaiseWorkflowStatus "...Rewriting datagram (from Archive) " & loRs!DatagramFormID & " with " & loRsChanges.RecordCount & " Changes to " & goSession.Site.GetSiteName(loRs!mwcSitesKey)
      Do While Not loRsChanges.EOF         '
         DoEvents
         sTableName = ""
         lorsChangeTable.Filter = adFilterNone
         lorsChangeTable.Filter = "ID=" & loRsChanges.Fields(RS_CL_mwrChangeTableKey).value
         If Not (lorsChangeTable.BOF And lorsChangeTable.EOF) Then
            sTableName = lorsChangeTable!TableName
         End If
         lorsChangeTable.Filter = adFilterNone
         Select Case loRsChanges.Fields(RS_CL_ChangeType).value
            Case "D"
               ' Delete Transaction
               With loform.DataForm.AddRecord
                  .Add "OPERATION", "D", "OPERATION"
                  .Add "TABLE", sTableName, "TABLE"
                  .Add "ID", loRsChanges.Fields(RS_CL_RecordKey).value, "ID"
               End With
            Case "A"
               ' Add Transaction
               Set loRec = loform.DataForm.AddRecord
               loRec.Add "OPERATION", "A", "OPERATION"
               loRec.Add "TABLE", sTableName, "TABLE"
               If Not WriteAuditChangeColumnsFromArchive(dbCon, sArchiveTableChildName, loRec, loRsChanges.Fields(RS_CL_ID).value) Then
                  CloseRecordset loRs
                  CloseRecordset loRsChanges
                  CloseRecordset lorsChangeTable
                  goSession.RaiseWorkflowStatus "Error Creating Batch Datagrams. Process cancelled."
                  Exit Function
               End If
            Case "M"
               ' Modify Transaction
               Set loRec = loform.DataForm.AddRecord
               loRec.Add "OPERATION", "M", "OPERATION"
               loRec.Add "TABLE", sTableName, "TABLE"
               loRec.Add "ID", loRsChanges.Fields(RS_CL_RecordKey).value, "ID"
               If Not WriteAuditChangeColumnsFromArchive(dbCon, sArchiveTableChildName, loRec, loRsChanges.Fields(RS_CL_ID).value) Then
                  CloseRecordset loRs
                  CloseRecordset loRsChanges
                  CloseRecordset lorsChangeTable
                  goSession.RaiseWorkflowStatus "Error Creating Batch Datagrams. Process cancelled."
                  Exit Function
               End If
               
            Case "S"
               ' Schema Transaction
               Set loRec = loform.DataForm.AddRecord
               
               loRec.Add "OPERATION", "S", "OPERATION"
               loRec.Add "TABLE", sTableName, "TABLE"
               If Not WriteAuditChangeColumnsFromArchive(dbCon, sArchiveTableChildName, loRec, loRsChanges.Fields(RS_CL_ID).value) Then
                  CloseRecordset loRs
                  CloseRecordset loRsChanges
                  CloseRecordset lorsChangeTable
                  goSession.RaiseWorkflowStatus "Error Creating Batch Datagrams. Process cancelled."
                  Exit Function
               End If
               
            Case "L"
               ' send Log Transaction email (like Schema Transaction)
               Set loRec = loform.DataForm.AddRecord
               
               loRec.Add "OPERATION", "L", "OPERATION"
               loRec.Add "TABLE", sTableName, "TABLE"
               If Not WriteAuditChangeColumnsFromArchive(dbCon, sArchiveTableChildName, loRec, loRsChanges.Fields(RS_CL_ID).value) Then
                  CloseRecordset loRs
                  CloseRecordset loRsChanges
                  CloseRecordset lorsChangeTable
                  goSession.RaiseWorkflowStatus "Error Creating Batch Datagrams. Process cancelled."
                  Exit Function
               End If
               
            Case Else
               goSession.RaiseWorkflowStatus "Warning in mwReplicateBatchWork. invalid change type code. Key: " & loRsChanges.Fields(RS_CL_ID).value
         End Select
         '
         loRsChanges.MoveNext
      Loop
      '
      ' Create Integrity Trailor Record
      '
      With loform.DataForm.AddRecord
         .Add "Operation", "I"
         .Add "COMMENT", "Batch Integrity Record"
         .Add "SiteID", goSession.Site.SiteID, "SiteID"
         If goSession.Site.SiteType <> SITE_TYPE_SHORE Then
            ' Ship to shore
            .Add "TargetSiteID", goSession.Site.TargetReplicateSiteID, "TargetSiteID"
            .Add "TargetSiteKey", goSession.Site.GetSiteKey(goSession.Site.TargetReplicateSiteID), "TargetSiteKey"
         Else
            ' Shore to ship
            .Add "TargetSiteID", goSession.Site.GetSiteID(loRs!mwcSitesKey), "TargetSiteID"
            .Add "TargetSiteKey", loRs!mwcSitesKey, "TargetSiteKey"
         End If
         .Add "BatchType", loRs!mwrBatchTypeKey, "BatchType"
         .Add "BatchNumber", loRs!BatchNumber, "BatchNumber"
         ' added by TJM
         If Not IsNull(loRs!mwrBatchLogOutboundFleetKey) Then
            .Add "FleetBatchNumber", loRs!mwrBatchLogOutboundFleetKey, "FleetBatchNumber"
         Else
            .Add "FleetBatchNumber", 0, "FleetBatchNumber"
         End If
         .Add "BatchKey", CStr(loRs!ID), "BatchKey"
         .Add "TotalRecords", loform.DataForm.Records.Count
      End With
      
      '
      '
      DoEvents
      If Not loform.DataForm.WriteDataForm(goSession.Workflow.ArchiveOkZipFile) Then
         goSession.RaiseError "Error in mwReplicateBatchWork.RemakeTransmissionFileFromArchive, WriteDataForm Failed. Operation Cancelled."
         CloseRecordset loRsChanges
         CloseRecordset loRs
         CloseRecordset lorsChangeTable
         loform.TearDown
         RemakeTransmissionFileFromArchive = False
         Exit Function
      End If
   End If
   CloseRecordset loRsChanges
   CloseRecordset loRs
   CloseRecordset lorsChangeTable
   If sAltDBConnectionString <> "" Then
      Set dbCon = Nothing
   End If
   '
   RemakeTransmissionFileFromArchive = True
   Exit Function
FunctionError:
'Resume Next
   goSession.RaisePublicError "General Error in mwReplicateBatchWork.RemakeTransmissionFileFromArchive. ", Err.Number, Err.Description
End Function


Private Function WriteAuditChangeColumnsFromArchive(dbCon As Connection, sArchiveTableChildName As String, oRec As ConfigKeys, ChangeLogKey As Long) As Boolean
   Dim ssql As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   ssql = "SELECT *  from " & sArchiveTableChildName & " where mwrChangeLogKey=" & ChangeLogKey & " ORDER BY ID"
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open ssql, dbCon, adOpenForwardOnly, adLockReadOnly
   Do While Not loRs.EOF
      If Not IsNull(loRs!NewValue) Then
         AddRecordField oRec, loRs!ColumnName, loRs!NewValue, loRs!ColumnDataType
      ElseIf Not IsNull(loRs!NewValueLarge) Then
         AddRecordField oRec, loRs!ColumnName, loRs!NewValueLarge, loRs!ColumnDataType
      Else
         goSession.RaiseWorkflowStatus "Warning in mwReplicateBatchWork.WriteAuditChangeColumnsFromArchive, no value to replicate. Key: " & loRs!ID
      End If
      loRs.MoveNext
   Loop
   loRs.Close
   Set loRs = Nothing
   WriteAuditChangeColumnsFromArchive = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwReplicateBatchWork.WriteAuditChangeColumnsFromArchive. ", Err.Number, Err.Description
End Function

Public Function ResendFileCabinetChanges(nBatchKey As Long, sFullFilename As String, nBatchType As Long, nBatchNo As Long) As Boolean
   Dim loRsFAB As Recordset
   Dim loRs As Recordset
   Dim ssql As String
   
   On Error GoTo FunctionError
   ssql = "select * from smfilecabinet where mwrBatchLogOutboundkey=" & nBatchKey
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open ssql, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      goSession.RaisePublicError "Error in mwReplicateBatchWork.ResendFileCabinetChanges, File Cabinet Record Missing for mwrBatchLogOutboundkey " & nBatchKey
      CloseRecordset loRs
      ResendFileCabinetChanges = False
      Exit Function
   End If

   Set loRsFAB = New Recordset
   With loRsFAB.Fields
      .Append "ID", adInteger, 4
      .Append "mwcSitesKey", adInteger, 4
      .Append "BatchType", adInteger, 4
      .Append "BatchKey", adInteger, 4
      .Append "BatchNumber", adInteger, 4
      .Append "DateMod", adDate
      .Append "FileBlob", adLongVarBinary, LONG_BIN_DEFINED_SIZE, adFldLong
   End With
   
   loRsFAB.Open
   Do While Not loRs.EOF
      loRsFAB.AddNew
      loRsFAB!ID = loRs!ID
      loRsFAB!mwcSitesKey = loRs!mwcSitesKey
      loRsFAB!BatchType = nBatchType
      loRsFAB!BatchKey = nBatchKey
      loRsFAB!BatchNumber = nBatchNo
      loRsFAB!DateMod = loRs!DateMod
      loRsFAB!FileBlob = loRs!FileBlob
      
      loRsFAB.Update
      loRs.MoveNext
   Loop
   
   loRsFAB.Save sFullFilename, adPersistADTG

   CloseRecordset loRs
   CloseRecordset loRsFAB
   ResendFileCabinetChanges = True
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwReplicateBatchWork.ResendFileCabinetChanges: ", Err.Number, Err.Description
   CloseRecordset loRs
   CloseRecordset loRsFAB
End Function


   ' Find ChangeLog records with NULL mwcSitesKeys because .NET doesn't fill it in.
   
Private Sub FixChangeLogSiteKeys()
   Dim ssql As String
   Dim loRs As Recordset
   Dim nSitesKey As Long
   Dim nNumUpdated As Long
   
   On Error GoTo FunctionError

   goSession.RaiseWorkflowStatus "FixChangeLogSiteKeys Looking for NULL mwcSitesKey entries in mwrChangeLog"

   goSession.MemoryLogIt "FixChangeLogSiteKeys Start"

   nNumUpdated = 0
   
   ssql = "SELECT * FROM mwrChangeLog WHERE mwcSitesKey Is Null AND ChangeType IN('A', 'L', 'M')"
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open ssql, goCon, adOpenDynamic, adLockOptimistic

   If IsRecordLoaded(loRs) Then
      goSession.RaiseWorkflowStatus "FixChangeLogSiteKeys Found " & loRs.RecordCount & " NULL mwcSitesKey entries."
   
      Do While Not loRs.EOF
         nSitesKey = GetSitesKey(ZeroNull(loRs!mwrChangeTableKey), ZeroNull(loRs!RecordKey))
         
         ' If we get a non-zero SitesKey then stuff it in this mwrChangeLog record.
         If nSitesKey <> 0 Then
            loRs!mwcSitesKey = nSitesKey
            loRs.Update
            nNumUpdated = nNumUpdated + 1
         End If
         loRs.MoveNext
      Loop
   End If
   
   CloseRecordset loRs
   
   goSession.RaiseWorkflowStatus "FixChangeLogSiteKeys Completed, " & nNumUpdated & " records updated"

   goSession.MemoryLogIt "FixChangeLogSiteKeys End"
   
   Exit Sub
FunctionError:
   goSession.RaisePublicError "General Error in mwIntegrate.WFSchedule.FixChangeLogSiteKeys. ", Err.Number, Err.Description

End Sub

' Get the underlying record for this mwrChangeLog entry and extract the mwcSitesKey if there is one.
' If any errors happen just return a zero

Private Function GetSitesKey(mwrChangeTableKey As Long, RecordKey As Long) As Long
   Dim ssql As String
   Dim loRs As Recordset
   Dim loRs2 As Recordset
   
   On Error GoTo FunctionError

   GetSitesKey = 0   ' Zero means no mwcSitesKey found
   
   ssql = "SELECT * FROM mwrChangeTable WHERE ID = " & mwrChangeTableKey
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open ssql, goCon, adOpenForwardOnly, adLockReadOnly

   If IsRecordLoaded(loRs) Then
      If loRs!mwrBatchTypeKey = 100 And BlankNull(loRs!TableName) <> "" Then
      
         ssql = "SELECT mwcSitesKey FROM " & loRs!TableName & " WHERE ID = " & RecordKey
         Set loRs2 = New Recordset
         loRs2.CursorLocation = adUseClient
         loRs2.Open ssql, goCon, adOpenForwardOnly, adLockReadOnly
      
         If IsRecordLoaded(loRs2) Then
            GetSitesKey = ZeroNull(loRs2!mwcSitesKey)
         End If
         CloseRecordset loRs2
      End If
   End If
   
   CloseRecordset loRs
   
   Exit Function
FunctionError:
   CloseRecordset loRs
   CloseRecordset loRs2
   GetSitesKey = 0

End Function


