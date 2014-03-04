Attribute VB_Name = "modStaging"
Option Explicit

Public gAddEventFactSiteKey As Long

Public gIsAllowWin9xDelay As Boolean
Public gWin9xMilliseconds As Long
Public goSession As Session
Public gIsSqlServer As Boolean
Public gIsOracle As Boolean
Public gIsAccess As Boolean
Public gAppPath As String

Public goCon As Connection
Public goConShape As Connection
Public goConBlob As Connection

Public Const ENCRYPT_PSWD = "Gray" & "bar" & "327"
Public gDBConnectString As String
Public gDbShapeConnectString As String

' 12/20/2004 Cheap solution...
Public gIsLeaveTransFile As Boolean

Public Const FORM_HEADER_STATUS_NULL = 0
Public Const FORM_HEADER_STATUS_PENDING = 1
Public Const FORM_HEADER_STATUS_SUBMITTED = 2
Public Const FORM_STATUS_STATUS_ARCHIVED = 3
Public Const FORM_STATUS_STATUS_SIGNEDOUT = 4

'
' License Options
'
Public Const LIC_01_Crewing = 1
Public Const LIC_02_Vessel_Reporting = 2
Public Const LIC_03_Safety_Management = 3
Public Const LIC_04_Document_Control = 4
Public Const LIC_05_Warranty_Claims = 5
Public Const LIC_06_ShipWorks_Equipment = 6
Public Const LIC_07_ShipWorks_Drydock = 7
Public Const LIC_08_ShipWorks_Requisitioning = 8
Public Const LIC_09_Maintenance = 9
Public Const LIC_10_MarineAssurance = 10


'Option2=Edit Management Lists!
'Option3=Safety Management!
'Option4=Data Logger!
'Option5=SDK _ Developer's Kit!
'Option6 = ShipWorks - Equipment
'Option7=Dry dock Module
'Option8 = ShipWorks - Ordering!
'Option9 = ShipWorks - Maintenance



'
Public goRepForm As frmCrystalViewer

Public Const MWRT_mwEventHistoryLog = 959
Public Const MWRT_mwEventLinkLog = 960
Public Const MWRT_mwEventFactLogSN = 961
Public Const MWRT_mwEventFormLog = 962
Public Const MWRT_mwwfFormTemplateEmailList = 963
Public Const MWRT_mwEventPropertyLog = 964
Public Const MWRT_mwAppSequence = 965
Public Const MWRT_mwEventLog = 966

Public Const MWRT_mwAlertLog = 969
Public Const MWRT_mwAlertLogStatus = 970
Public Const MWRT_mwAlertType = 971
Public Const MWRT_mwAlertEvents = 972
Public Const MWRT_mwAlertDist = 973

Public Const MWRT_mwBlobFile = 981

Public Const MWRT_mwFormStatus = 600
Public Const MWRT_mwFormCabinet = 601
Public Const MWRT_mwFormFolders = 602
Public Const MWRT_mwFormHeader = 603
Public Const MWRT_mwFormTemplateRev = 604
Public Const MWRT_mwFormFolderContents = 605
Public Const MWRT_mwFormDetail = 606

Public Const MWRT_mwcSchReplicate = 20001
Public Const MWRT_mwcSchReplicateShip = 30043

Public Const MWRT_smManualChapter = 302
Public Const MWRT_mwwfFormTemplate = 907

Public Const DATAGRAM_DATE_FORMAT = "YYYY-MM-DD hh:nn:ss"
Public Const DATAGRAM_SHORT_DATE_FORMAT = "YYYY-MM-DD"
Public Const DATAGRAM_DECIMAL_FORMAT = "."

Public Const SITE_TYPE_SHIP = 1
Public Const SITE_TYPE_SHORE = 2

Public Const ENCRYPT_PASSWORD = "eGw4918"

Public Const MW_EVENT_WORKFLOW_AGENT = 908
Public Const MW_EVENT_Certificate_Ship = 20401
Public Const MW_EVENT_Eqpt_History = 103
Public Const MW_EVENT_Requisiton_Header = 110
Public Const SW_EVENT_HISTORY = 103
Public Const MW_EVENT_VRS_VOYEVENT = 20301
Public Const MW_EVENT_PMS_CHANGE_REQUEST = 80001
Public Const MW_EVENTTYPE_ADMIN_CFG_SWITCHES = 999

Public Const MW_ALERT_STATUS_SENT = 1
Public Const MW_ALERT_STATUS_READ = 2
Public Const MW_ALERT_STATUS_REPLIED = 3
Public Const MW_ALERT_STATUS_CLOSED = 4

Public Const MW_ALERT_TYPE_USER = 1
Public Const MW_ALERT_TYPE_SYSTEM = 2
Public Const MW_ALERT_TYPE_ADMIN = 3

' frmCfgBrowser mwrChangeTable
   Public Const PV_CT_ID = 0
   Public Const PV_CT_TableName = 1
   Public Const PV_CT_IsActive = 2
   Public Const PV_CT_mwrBatchTypeKey = 3
   Public Const PV_CT_TableDescription = 4
   Public Const PV_CT_SaveAuditLogs = 5
   
' grid save to registry constant
   Public Const BASE_REG = "Software\Maritime Systems Inc\SessionSettings\"

Public Const SW_EV_USERS = 824

Public Const MWRT_smOccurrence = 320
Public Const MWRT_smOccurrenceFinding = 324
Public Const MWRT_smOccurrenceInjury = 331

Public Const VRS_FACT_TYPE_CANAL = 1
Public Const VRS_FACT_TYPE_BUNKERING = 2
Public Const VRS_FACT_TYPE_CARGO_LOAD = 3
Public Const VRS_FACT_TYPE_CARGO_DISCH = 4
Public Const VRS_FACT_TYPE_CARGO_BOTH = 5

'
' Valid Container codes for Desktop INI File
'
Public Const MW_CONTAINER_OFFICE = "OFF"
Public Const MW_CONTAINER_VBDLL = "VBDLL"
Public Const MW_CONTAINER_MW_EVENTS = "MW_EVENT"
Public Const MW_CONTAINER_WF_EVENTS = "WF_EVENT"
Public Const MW_CONTAINER_WF_EXPLORE = "WF_EXPLORE"
Public Const MW_CONTAINER_MANAGEMENT_LIST = "ML"
Public Const MW_CONTAINER_VISIO = "V"
Public Const MW_CONTAINER_CRYSTAL = "R"
Public Const MW_CONTAINER_MANUAL = "H"
Public Const MW_CONTAINER_FORM = "F"
Public Const MW_CONTAINER_FLUKE = "L"
Public Const MW_CONTAINER_USER = "U"
Public Const MW_CONTAINER_SHELL_OUT = "S"
Public Const MW_CONTAINER_MAILBOX_AGENT = "MA"
Public Const MW_CONTAINER_FOLDER_AGENT = "FA"
Public Const MW_CONTAINER_SCRIPT_AGENT = "SA"
Public Const MW_CONTAINER_CRYSTAL_AGENT = "CA"
Public Const MW_CONTAINER_MMS_AGENT = "MMSA"
Public Const MW_CONTAINER_SMS_ROUTER = "SMS"
Public Const MW_CONTAINER_E2S_ROUTER = "E2S"

'
' Electronic Safety Platform Containers... 5/18/2001 ms
'
Public Const MW_CONTAINER_ESP_DISTRIBUTION = "DD"
Public Const MW_CONTAINER_ESP_RECEIPT = "DR"
Public Const MW_CONTAINER_ESP_LISTS = "LI"

