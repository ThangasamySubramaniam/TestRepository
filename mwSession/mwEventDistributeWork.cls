VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwEventDistributeWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' mwEventDistributeWork - Distribute Events work class
' 17 May 2003 ms
'

Option Explicit


Public Function AddEventSites(EventType As Long, EventDetailKey As Long) As Boolean
   Dim loRsSites As Recordset
   Dim strSQL As String
   On Error GoTo FunctionError
   '
   '
   ' Start Transaction
   '
   '
   ' Add Site Records...
   '
   loRsSites.MoveFirst
   Do While Not loRsSites.EOF
      strSQL = "Insert into mwEventDistributeSites (pkgPackageKey,mwcSitesKey,SiteName,NoTimesSent,SendUpdate,ConfirmedReceipt) values (" & _
        moRS!ID & COMMAQ & loRsSites!ID & QCOMMAQ & loRsSites!SiteName & "',0,0,0)"
      
      goCon.Execute strSQL
      loRsSites.MoveNext
   Loop
   ugSites.Update
   '
   Set loVaWork = New mwEventFactsWork
   loVaWork.AutoCreate moInterface.EventType, moInterface.EventDetailKey
   
   '
   ' Commit Transaction
   '
   goCon.CommitTrans
   
   LoadPackageSite
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSafety.frmRemotePackages.AddRecord. ", err.Number, err.Description
   '
   ' Rollback on error
   '
   goCon.RollbackTrans
End Function


