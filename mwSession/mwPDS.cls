VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwPDS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'
' Package Delivery System
'
' 15 August 2005 T. McSally

Option Explicit

'10 = CREATE TABLE pdsPackage (
'     ID                      long not null CONSTRAINT pdsPackage_PK PRIMARY KEY,
'     OriginatingSite         long,
'     DeliverToSite           long,
'     pdsStatusKey            long,
'     CreatedBy               long,
'     pdsPriority             long,
'     CreatedDate             datetime,
'     DeliveredDate           datetime,
'     DeliveryCompleteDate    datetime,
'     pdsDeliveryMethodKey    long)
'
'
'11 = CREATE TABLE pdsPackageItem (
'     ID                      long not null CONSTRAINT pdsPackageItem_PK PRIMARY KEY,
'     pdsPackageKey           long,
'     pdsItemTypeKey          long,
'     pdsStatusKey            long,
'     Description             varchar(50),
'     pdsPriority             long,
'     ReceivedBy              long,
'     DeliveryCompleteDate    datetime,
'     SourceFileName          varchar(64),
'     TargetPlaceID           varchar(50),
'     SourceBLOBKey           long,
'     Reference               varchar(50),
'     ManualChapterKey        long,
'     ManualRevisionNo        long,
'     ManualRevisionDate      datetime)
'
'
'20 = CREATE TABLE pdsItem (
'     ID                      long not null CONSTRAINT pdsItem_PK PRIMARY KEY,
'     pdsItemTypeKey          long,
'     pdsStatusKey            long,
'     Description             varchar(50),
'     pdsPriority             long,
'     RequestedBy             long,
'     RequestedDate           datetime,
'     DeliveryCompleteDate    datetime,
'     SourceFileName          varchar(64),
'     SourceFileLocation      varchar(200),
'     StagingFileLocation     varchar(200),
'     TargetPlaceID           varchar(50),
'     SourceBLOBKey           long,
'     Reference               varchar(50),
'     ManualChapterKey        long,
'     ManualRevisionNo        long,
'     ManualRevisionDate      datetime)
'
'
'21 = CREATE TABLE pdsItemSite (
'     ID                      long not null CONSTRAINT pdsItemSite_PK PRIMARY KEY,
'     pdsItemKey              long,
'     DeliverToSite           long,
'     pdsStatusKey            long,
'     pdsPriority             long,
'     pdsPackageKey           long,
'     pdsPackageItemKey       long,
'     PackagedBy              long,
'     PackagedDate            datetime,
'     DeliveredDate           datetime,
'     pdsDeliveryMethodKey    long)

Private WithEvents moItemSiteRs As Recordset
Attribute moItemSiteRs.VB_VarHelpID = -1
Private WithEvents moItemRs As Recordset
Attribute moItemRs.VB_VarHelpID = -1
Private WithEvents moPackageItemRs As Recordset
Attribute moPackageItemRs.VB_VarHelpID = -1
Private WithEvents moPackageRs As Recordset
Attribute moPackageRs.VB_VarHelpID = -1

Public Enum PDS_Type
   PDS_MWBLOB = 10
   PDS_FCBLOB = 11
   PDS_FILE = 30
   PDS_MANCHAPTER = 31
End Enum

'02a = INSERT INTO pdsItemType (ID, Description ) VALUES (10,'BLOB Item')
'02b = INSERT INTO pdsItemType (ID, Description ) VALUES (11,'File Cabinet Item')
'02c = INSERT INTO pdsItemType (ID, Description ) VALUES (30,'File Item')
'02d = INSERT INTO pdsItemType (ID, Description ) VALUES (31,'Manual Chapter Item')

Public Enum PDS_Priority
   PDS_Prio_None = 0
   PDS_Prio_Hi = 1
   PDS_Prio_Med = 2
   PDS_Prio_Low = 3
End Enum

'04a = INSERT INTO pdsPriority (ID, Description ) VALUES (10,'High')
'04b = INSERT INTO pdsPriority (ID, Description ) VALUES (20,'Medium')
'04c = INSERT INTO pdsPriority (ID, Description ) VALUES (30,'Low')

Public Enum PDS_Status
   PDS_Stat_Init = 0
   PDS_Stat_RdyDel = 10
   PDS_Stat_AsgnPkg = 20
   PDS_Stat_RdyShip = 30
   PDS_Stat_Shipped = 40
   PDS_Stat_Rcvd = 50
   PDS_Stat_Delivered = 60
   PDS_Stat_Complete = 70
   PDS_Stat_Cancelled = 80
End Enum


'01A = INSERT INTO pdsStatus (ID, Description ) VALUES (00,'Initial Request for Pickup')
'01B = INSERT INTO pdsStatus (ID, Description ) VALUES (10,'Ready for Delivery')
'01C = INSERT INTO pdsStatus (ID, Description ) VALUES (20,'Assigned to Package')
'01D = INSERT INTO pdsStatus (ID, Description ) VALUES (30,'Package Waiting for Shipping')
'01E = INSERT INTO pdsStatus (ID, Description ) VALUES (40,'Package Shipped')
'01F = INSERT INTO pdsStatus (ID, Description ) VALUES (50,'Package Received')
'01G = INSERT INTO pdsStatus (ID, Description ) VALUES (60,'Item Delivered')
'01H = INSERT INTO pdsStatus (ID, Description ) VALUES (70,'Delivery Complete')
'01I = INSERT INTO pdsStatus (ID, Description ) VALUES (80,'Delivery Cancelled')

Public Enum PDS_SendBy
   PDS_SendByEmail = 1
   PDS_SendByMedia = 2
   PDS_SendByCDROM = 3
   PDS_SendByCopy = 4
End Enum

'03A = INSERT INTO pdsDeliveryMethod (ID, Description ) VALUES (1,'EMail')
'03B = INSERT INTO pdsDeliveryMethod (ID, Description ) VALUES (2,'Media')
'03C = INSERT INTO pdsDeliveryMethod (ID, Description ) VALUES (3,'CDROM')
'03D = INSERT INTO pdsDeliveryMethod (ID, Description ) VALUES (4,'Direct Copy')


Private Const MWRT_pdsStatus = 805
Private Const MWRT_pdsItemType = 806
Private Const MWRT_pdsDeliveryMethod = 807
Private Const MWRT_pdsPriority = 808
Private Const MWRT_pdsPackage = 809
Private Const MWRT_pdsPackageItem = 810
Private Const MWRT_pdsItem = 811
Private Const MWRT_pdsItemSite = 812

' Working Objects

Private moParent As Session


Friend Property Set Parent(oParent As mwSession.Session)
   If moParent Is Nothing Then
      Set moParent = oParent
   End If
End Property

Friend Property Get Parent() As mwSession.Session
   Set Parent = moParent
End Property


'Public Property Let A(ByVal Data As String)
'    mA = Data
'End Property

'Public Property Get A() As String
'    A = mA
'End Property

Private Sub Class_Initialize()
' Initialize code goes here
   On Error GoTo SubError
   
   Exit Sub
SubError:
   goSession.RaiseError "Error in mwSession.mwPDS.Class_Initialize.", Err.Number, Err.Description
End Sub

Private Sub Class_Terminate()
' Cleanup code goes here
   On Error GoTo SubError
   
   Exit Sub
SubError:
   goSession.RaiseError "Error in mwSession.mwPDS.Class_Terminate.", Err.Number, Err.Description
End Sub



Public Function PostItem(ItemType As PDS_Type, ItemPriority As PDS_Priority, ItemBLOB As Long, ItemFileName As String, Description As String, Optional TargetPlaceID As String, Optional DistFleet As Long) As Long
   
   On Error GoTo FunctionError
   
   If ItemType = PDS_MWBLOB Then
      PostItem = Post_MWBLOBItem(ItemPriority, ItemBLOB, Description, DistFleet)
      
   ElseIf ItemType = PDS_FCBLOB Then
      PostItem = Post_FCBLOBItem(ItemPriority, ItemBLOB, Description, DistFleet)
      
   ElseIf ItemType = PDS_FILE Then
      PostItem = Post_FileItem(ItemPriority, ItemFileName, Description, TargetPlaceID, DistFleet)
   
   Else
      goSession.RaiseError "Error in mwSession.mwPDS.Post_FileItem, Invalid PDS_Type " & ItemType, Err.Number, Err.Description
      PostItem = -1
   End If
   
   Exit Function
FunctionError:
   goSession.RaiseError "Error in mwSession.mwPDS.PostItem.", Err.Number, Err.Description
   PostItem = -1
End Function

'20 = CREATE TABLE pdsItem (
'     ID                      long not null CONSTRAINT pdsItem_PK PRIMARY KEY,
'     pdsItemTypeKey          long,
'     pdsStatusKey            long,
'     Description             varchar(50),
'     pdsPriority             long,
'     RequestedBy             long,
'     RequestedDate           datetime,
'     DeliveryCompleteDate    datetime,
'     SourceFileName          varchar(64),
'     SourceFileLocation      varchar(200),
'     StagingFileLocation     varchar(200),
'     TargetPlaceID           varchar(50),
'     SourceBLOBKey           long,
'     Reference               varchar(50),
'     ManualChapterKey        long,
'     ManualRevisionNo        long,
'     ManualRevisionDate      datetime)
'
'


Public Function Post_MWBLOBItem(ItemPriority As PDS_Priority, ItemBLOB As Long, Description As String, Optional DistFleet As Long) As Long
   On Error GoTo FunctionError
   
   Set moItemRs = FetchPDSItemRs(-1)
   
   Post_MWBLOBItem = goSession.MakePK("pdsItem")
   
   moItemRs.AddNew
   moItemRs!ID = Post_MWBLOBItem
   moItemRs!pdsItemTypeKey = PDS_MWBLOB
   moItemRs!pdsStatusKey = PDS_Stat_Init
   moItemRs!Description = Description
   moItemRs!pdsPriority = ItemPriority
   moItemRs!RequestedBy = goSession.User.userkey
   moItemRs!DeliveryCompleteDate = Null
   moItemRs!SourceFilename = Null
   moItemRs!SourceFileLocation = Null
   moItemRs!StagingFileLocation = Null
   moItemRs!TargetPlaceID = Null
   moItemRs!SourceBLOBKey = ItemBLOB
   moItemRs!Reference = Null
   moItemRs!ManualChapterKey = Null
   moItemRs!ManualRevisionNo = Null
   moItemRs!ManualRevisionDate = Null
      
   moItemRs.Update
   
   If DistFleet > 0 Then
      SendItemToFleet Post_MWBLOBItem, DistFleet
   End If
   
   Exit Function
FunctionError:
   goSession.RaiseError "Error in mwSession.mwPDS.Post_MWBLOBItem.", Err.Number, Err.Description
   Post_MWBLOBItem = -1
End Function



Public Function Post_FCBLOBItem(ItemPriority As PDS_Priority, ItemBLOB As Long, Description As String, Optional DistFleet As Long) As Long
   On Error GoTo FunctionError
   
   Set moItemRs = FetchPDSItemRs(-1)
   
   Post_FCBLOBItem = goSession.MakePK("pdsItem")
   
   moItemRs.AddNew
   moItemRs!ID = Post_FCBLOBItem
   moItemRs!pdsItemTypeKey = PDS_FCBLOB
   moItemRs!pdsStatusKey = PDS_Stat_Init
   moItemRs!Description = Description
   moItemRs!pdsPriority = ItemPriority
   moItemRs!RequestedBy = goSession.User.userkey
   moItemRs!DeliveryCompleteDate = Null
   moItemRs!SourceFilename = Null
   moItemRs!SourceFileLocation = Null
   moItemRs!StagingFileLocation = Null
   moItemRs!TargetPlaceID = Null
   moItemRs!SourceBLOBKey = ItemBLOB
   moItemRs!Reference = Null
   moItemRs!ManualChapterKey = Null
   moItemRs!ManualRevisionNo = Null
   moItemRs!ManualRevisionDate = Null
      
   moItemRs.Update
   
   If DistFleet > 0 Then
      SendItemToFleet Post_FCBLOBItem, DistFleet
   End If
   
   Exit Function
FunctionError:
   goSession.RaiseError "Error in mwSession.mwPDS.Post_FCBLOBItem.", Err.Number, Err.Description
   Post_FCBLOBItem = -1
End Function



Public Function Post_FileItem(ItemPriority As PDS_Priority, ItemFileName As String, Description As String, TargetPlaceID As String, Optional DistFleet As Long) As Long
   Dim fso As FileSystemObject
   Dim SourcePath As String
   Dim FName As String
   
   On Error GoTo FunctionError
   
   If Not fso.FileExists(ItemFileName) Then
      goSession.RaiseError "Error in mwSession.mwPDS.Post_FileItem: File does not exist - " & ItemFileName, Err.Number, Err.Description
      Post_FileItem = -1
      Exit Function
   End If
   
   SourcePath = fso.GetParentFolderName(ItemFileName)
   FName = fso.GetFileName(ItemFileName)
   
   Post_FileItem = goSession.MakePK("pdsItem")
   
   Set moItemRs = FetchPDSItemRs(-1)
   
   moItemRs.AddNew
   moItemRs!ID = Post_FileItem
   moItemRs!pdsItemTypeKey = PDS_FILE
   moItemRs!pdsStatusKey = PDS_Stat_Init
   moItemRs!Description = Description
   moItemRs!pdsPriority = ItemPriority
   moItemRs!RequestedBy = goSession.User.userkey
   moItemRs!DeliveryCompleteDate = Null
   moItemRs!SourceFilename = Null
   moItemRs!SourceFileLocation = Null
   moItemRs!StagingFileLocation = Null
   moItemRs!TargetPlaceID = Null
   moItemRs!SourceBLOBKey = Null
   moItemRs!Reference = Null
   moItemRs!ManualChapterKey = Null
   moItemRs!ManualRevisionNo = Null
   moItemRs!ManualRevisionDate = Null
      
   moItemRs.Update
   
   If DistFleet > 0 Then
      SendItemToFleet Post_FileItem, DistFleet
   End If
   
   Exit Function
FunctionError:
   goSession.RaiseError "Error in mwSession.mwPDS.Post_FileItem.", Err.Number, Err.Description
   Post_FileItem = -1
End Function



Public Function SendItemToSite(pdsItemKey As Long, mwcSitesKey As Long) As Boolean
   On Error GoTo FunctionError
   
   Exit Function
FunctionError:
   goSession.RaiseError "Error in mwSession.mwPDS.SendItemToSite.", Err.Number, Err.Description
   SendItemToSite = False
End Function



Public Function SendItemToFleet(pdsItemKey As Long, mwcFleetsKey As Long) As Boolean
   On Error GoTo FunctionError
   
   Exit Function
FunctionError:
   goSession.RaiseError "Error in mwSession.mwPDS.SendItemToFleet.", Err.Number, Err.Description
   SendItemToFleet = False
End Function



Public Function CreatePackage(mwcSitesKey As Long) As Long
   On Error GoTo FunctionError
   
   Exit Function
FunctionError:
   goSession.RaiseError "Error in mwSession.mwPDS.CreatePackage.", Err.Number, Err.Description
   CreatePackage = False
End Function



Public Function AddItemToPackage(pdsPackageKey As Long, pdsItemKey As Long) As Boolean
   On Error GoTo FunctionError
   
   Exit Function
FunctionError:
   goSession.RaiseError "Error in mwSession.mwPDS.AddItemToPackage.", Err.Number, Err.Description
   AddItemToPackage = False
End Function



Public Function RemoveItemFromPackage(pdsPackageKey As Long, pdsItemKey As Long) As Boolean
   On Error GoTo FunctionError
   
   Exit Function
FunctionError:
   goSession.RaiseError "Error in mwSession.mwPDS.RemoveItemFromPackage.", Err.Number, Err.Description
   RemoveItemFromPackage = False
End Function



Public Function DistributePackage(pdsPackageKey As Long, DistributionMethod As PDS_SendBy) As Boolean
   On Error GoTo FunctionError
   
   Exit Function
FunctionError:
   goSession.RaiseError "Error in mwSession.mwPDS.DistributePackage.", Err.Number, Err.Description
   DistributePackage = False
End Function



Public Function InstallPackage(PackageFile As String) As Boolean
   On Error GoTo FunctionError
   
   Exit Function
FunctionError:
   goSession.RaiseError "Error in mwSession.mwPDS.InstallPackage.", Err.Number, Err.Description
   InstallPackage = False
End Function



Public Function CreatePackageReceipt(pdsPackageKey As Long) As Boolean
   On Error GoTo FunctionError
   
   Exit Function
FunctionError:
   goSession.RaiseError "Error in mwSession.mwPDS.CreatePackageReceipt.", Err.Number, Err.Description
   CreatePackageReceipt = False
End Function



Public Function CreateItemReceipt(pdsPackageKey As Long, pdsItemKey As Long) As Boolean
   On Error GoTo FunctionError
   
   Exit Function
FunctionError:
   goSession.RaiseError "Error in mwSession.mwPDS.CreateItemReceipt.", Err.Number, Err.Description
   CreateItemReceipt = False
End Function


Private Sub moItemSiteRs_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   Static IsBeginAdd As Boolean
   Static IsBeginDelete As Boolean
   On Error GoTo SubError
   
   If adReason = adRsnAddNew Then
      IsBeginAdd = True
   ElseIf adReason = adRsnUpdate And IsBeginDelete Then
      IsBeginDelete = False
   ElseIf adReason = adRsnUpdate And IsBeginAdd Then
      '                                                 This is a FleetWide table, do NOT include a SitesKey
      goSession.ReplicateWork.LogAddChange MWRT_pdsItemSite, moItemSiteRs!ID, , moItemSiteRs.Fields
      IsBeginAdd = False
   ElseIf adReason = adRsnDelete Then
      goSession.ReplicateWork.LogDeleteChange MWRT_pdsItemSite, moItemSiteRs!ID
      IsBeginDelete = True
   ElseIf adReason <> adRsnFirstChange Then
      goSession.ReplicateWork.LogModifyChange MWRT_pdsItemSite, moItemSiteRs.Fields
   End If
   Exit Sub
SubError:
   goSession.RaisePublicError "General Error in mwSession.mwPDS.moItemSiteRs_WillChangeRecord. ", Err.Number, Err.Description

End Sub


Private Sub moItemRs_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   Static IsBeginAdd As Boolean
   Static IsBeginDelete As Boolean
   On Error GoTo SubError
   
   If adReason = adRsnAddNew Then
      IsBeginAdd = True
   ElseIf adReason = adRsnUpdate And IsBeginDelete Then
      IsBeginDelete = False
   ElseIf adReason = adRsnUpdate And IsBeginAdd Then
      '                                                 This is a FleetWide table, do NOT include a SitesKey
      goSession.ReplicateWork.LogAddChange MWRT_pdsItem, moItemRs!ID, , moItemRs.Fields
      IsBeginAdd = False
   ElseIf adReason = adRsnDelete Then
      goSession.ReplicateWork.LogDeleteChange MWRT_pdsItem, moItemRs!ID
      IsBeginDelete = True
   ElseIf adReason <> adRsnFirstChange Then
      goSession.ReplicateWork.LogModifyChange MWRT_pdsItem, moItemRs.Fields
   End If
   Exit Sub
SubError:
   goSession.RaisePublicError "General Error in mwSession.mwPDS.moItemRs_WillChangeRecord. ", Err.Number, Err.Description

End Sub

Private Sub moPackageItemRs_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   Static IsBeginAdd As Boolean
   Static IsBeginDelete As Boolean
   On Error GoTo SubError
   
   If adReason = adRsnAddNew Then
      IsBeginAdd = True
   ElseIf adReason = adRsnUpdate And IsBeginDelete Then
      IsBeginDelete = False
   ElseIf adReason = adRsnUpdate And IsBeginAdd Then
      '                                                 This is a Site Specific table, it MUST include a SitesKey
      goSession.ReplicateWork.LogAddChange MWRT_pdsPackageItem, moPackageItemRs!ID, moPackageItemRs!mwcSitesKey, moPackageItemRs.Fields
      IsBeginAdd = False
   ElseIf adReason = adRsnDelete Then
      goSession.ReplicateWork.LogDeleteChange MWRT_pdsPackageItem, moPackageItemRs!ID, moPackageItemRs!mwcSitesKey
      IsBeginDelete = True
   ElseIf adReason <> adRsnFirstChange Then
      goSession.ReplicateWork.LogModifyChange MWRT_pdsPackageItem, moPackageItemRs.Fields, moPackageItemRs!mwcSitesKey
   End If
   Exit Sub
SubError:
   goSession.RaisePublicError "General Error in mwSession.mwPDS.moPackageItemRs_WillChangeRecord. ", Err.Number, Err.Description

End Sub

Private Sub moPackageRs_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   Static IsBeginAdd As Boolean
   Static IsBeginDelete As Boolean
   On Error GoTo SubError
   
   If adReason = adRsnAddNew Then
      IsBeginAdd = True
   ElseIf adReason = adRsnUpdate And IsBeginDelete Then
      IsBeginDelete = False
   ElseIf adReason = adRsnUpdate And IsBeginAdd Then
      '                                                 This is a Site Specific table, it MUST include a SitesKey
      goSession.ReplicateWork.LogAddChange MWRT_pdsPackage, moPackageRs!ID, moPackageRs!mwcSitesKey, moPackageRs.Fields
      IsBeginAdd = False
   ElseIf adReason = adRsnDelete Then
      goSession.ReplicateWork.LogDeleteChange MWRT_pdsPackage, moPackageRs!ID, moPackageRs!mwcSitesKey
      IsBeginDelete = True
   ElseIf adReason <> adRsnFirstChange Then
      goSession.ReplicateWork.LogModifyChange MWRT_pdsPackage, moPackageRs.Fields, moPackageRs!mwcSitesKey
   End If
   Exit Sub
SubError:
   goSession.RaisePublicError "General Error in mwSession.mwPDS.moPackageRs_WillChangeRecord. ", Err.Number, Err.Description

End Sub

Private Function FetchPDSItemRs(pdsItemKey As Long) As Recordset
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   
   sSQL = "SELECT * FROM pdsItem WHERE ID=" & pdsItemKey
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   Set FetchPDSItemRs = loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwPDS.FetchPDSItemRs. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function

Private Function FetchPDSPackageRs(pdsPackageKey As Long) As Recordset
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   
   sSQL = "SELECT * FROM pdsPackage WHERE ID=" & pdsPackageKey
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   Set FetchPDSPackageRs = loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwPDS.FetchPDSPackageRs. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function

Private Function FetchPDSPackageItemRs(pdsPackageItemKey As Long) As Recordset
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   
   sSQL = "SELECT * FROM pdsPackageItem WHERE ID=" & pdsPackageItemKey
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   Set FetchPDSPackageItemRs = loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwPDS.FetchPDSPackageItemRs. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function

Private Function FetchPDSItemSiteRs(pdsItemSiteKey As Long) As Recordset
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   
   sSQL = "SELECT * FROM pdsItemSite WHERE ID=" & pdsItemSiteKey
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   Set FetchPDSItemSiteRs = loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwPDS.FetchPDSItemSiteRs. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function

