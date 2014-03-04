VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwMenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A menu item that appears under a menu that is displayed to the WFD user."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Session"
Option Explicit

Private mDisplayTitle As String
Private mItemTitle As String
Private mImageReference As String
Private mFormHandle As Object
Private mIsExists As Boolean
Private mIsValid As Boolean
Private mFormType As String
Private moFormTemplate As mwFormTemplate
Private mAutoLoad As Boolean
Private mMenuItemKey As Long

'Public Enum msWFDFormType
'    EXCEL_FORM = "E"
'    WORD_FORM = "W"
'    CUSTOM_FORM = "C"
'    CRYSTAL_REPORT = "R"
'End Type
Private mFormID As String
Private moExtendedKeys As ConfigKeys
Private moSes As Session

Public Property Set Session(oSession As Session)
   If moSes Is Nothing Then
      Set moSes = oSession
   End If
End Property

Public Property Get Session() As Session
   Set Session = moSes
End Property




Friend Property Set ExtendedKeys(ByVal vData As ConfigKeys)
   Set moExtendedKeys = vData
End Property

Friend Property Get ExtendedKeys() As ConfigKeys
   Set ExtendedKeys = moExtendedKeys
End Property

Public Property Let FormID(ByVal vData As String)
Attribute FormID.VB_Description = "Identifies the FormID; used for custom form initialization"
    mFormID = vData
End Property


Public Property Get FormID() As String
    FormID = mFormID
End Property

Public Property Let IsExists(ByVal vData As Boolean)
    mIsExists = vData
End Property


Public Property Get IsExists() As Boolean
    IsExists = mIsExists
End Property

Public Property Let IsValid(ByVal vData As Boolean)
    mIsValid = vData
End Property


Public Property Get IsValid() As Boolean
    IsValid = mIsValid
End Property



Public Property Let FormHandle(ByRef vData As Object)
    Set mFormHandle = vData
End Property



Public Property Get FormHandle() As Object
   Set FormHandle = mFormHandle
End Property

Public Property Let ImageReference(ByVal vData As String)
Attribute ImageReference.VB_Description = "Reference to the ImageList offset of the icon to be displayed."
    mImageReference = UCase(vData)
End Property


Public Property Get ImageReference() As String
   ImageReference = mImageReference
End Property



Public Property Let ItemTitle(ByVal vData As String)
    mItemTitle = vData
End Property


Public Property Get ItemTitle() As String
   If IsNull(ItemTitle) Then
      ItemTitle = ""
   Else
      ItemTitle = mItemTitle
   End If
End Property



Public Property Let DisplayTitle(ByVal vData As String)
    mDisplayTitle = vData
End Property


Public Property Get DisplayTitle() As String
   If IsNull(DisplayTitle) Then
      DisplayTitle = ""
   Else
      DisplayTitle = mDisplayTitle
   End If
End Property

Public Property Let FormType(ByVal Data As String)
    mFormType = Data
End Property


Public Property Get FormType() As String
    FormType = mFormType
End Property

Public Property Get FormTemplate() As mwFormTemplate
    Set FormTemplate = moFormTemplate
End Property

Public Property Let AutoLoad(ByVal Data As Boolean)
    mAutoLoad = Data
End Property

Public Property Get AutoLoad() As Boolean
    AutoLoad = mAutoLoad
End Property

Public Property Let MenuItemKey(ByVal Data As Long)
    mMenuItemKey = Data
End Property

Public Property Get MenuItemKey() As Long
    MenuItemKey = mMenuItemKey
End Property

'Public Property Let A(ByVal Data As String)
'    mA = Data
'End Property

'Public Property Get A() As String
'    A = mA
'End Property


Private Sub Class_Initialize()
   Set moFormTemplate = New mwFormTemplate
   Set moExtendedKeys = New ConfigKeys
   
End Sub

Private Sub Class_Terminate()
   Set moFormTemplate = Nothing
   Set moExtendedKeys = Nothing
End Sub



Public Function GetExtendedProperty(KeyName As String) As String
   On Error GoTo FunctionError
   If moExtendedKeys Is Nothing Then
      GetExtendedProperty = ""
      Exit Function
   End If
   GetExtendedProperty = moExtendedKeys(UCase(KeyName)).KeyValue
   Exit Function
FunctionError:
   GetExtendedProperty = ""
End Function

Public Function SetExtendedProperty(KeyName As String, KeyValue As String, Optional MIKey As Long) As Boolean
   Dim strSQL As String
   On Error GoTo FunctionError
   If moExtendedKeys Is Nothing Then
      SetExtendedProperty = False
      Exit Function
   End If
   If MIKey < 1 Then
      strSQL = "select * from mwcMenuItems where ID=" & mMenuItemKey
   Else
      strSQL = "select * from mwcMenuItems where ID=" & MIKey
   End If
   If moSes.UpdateExtendedKeyData(KeyName, KeyValue, strSQL) Then
      moExtendedKeys(UCase(KeyName)).KeyValue = KeyValue
   End If
   SetExtendedProperty = True
   Exit Function
FunctionError:
   moSes.RaiseError "General Error in mwMenuItem.SetExtendedProperty. Invalid Key ?: " & KeyName, err.Number, err.Description
   SetExtendedProperty = False
End Function

