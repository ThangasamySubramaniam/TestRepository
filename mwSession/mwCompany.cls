VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' mwCompany - Company Information for ShipNet Fleet
'
Option Explicit

Private mCompanyID As String
Private mCompanyName As String
Private mCompanyCode As String
'
' MTML "Company" specific properties
'
Private mMtmlInterchangeID As String
'Private mMtmlInterchangeSenderRecipient As String
'Private mMtmlInterchangeControlReference As String

Private moParent As mwSession.Session

Friend Property Set Parent(oParent As mwSession.Session)
   If moParent Is Nothing Then
      Set moParent = oParent
   End If
End Property

Friend Property Get Parent() As mwSession.Session
   'If moParent Is Nothing Then
   '   Set moParent = oParent
   'End If
   Set Parent = moParent
End Property




Public Property Let CompanyID(ByVal Data As String)
    mCompanyID = Data
End Property

Public Property Get CompanyID() As String
    CompanyID = mCompanyID
End Property

Public Property Let CompanyName(ByVal Data As String)
    mCompanyName = Data
End Property

Public Property Get CompanyName() As String
    CompanyName = mCompanyName
End Property

Public Property Let CompanyCode(ByVal Data As String)
    mCompanyCode = Data
End Property

Public Property Get CompanyCode() As String
    CompanyCode = mCompanyCode
End Property

'
'========== MTML Properties ====================
'
Public Property Let mtmlInterchangeID(ByVal Data As String)
    mMtmlInterchangeID = Data
End Property

Public Property Get mtmlInterchangeID() As String
    mtmlInterchangeID = mMtmlInterchangeID
End Property

'Public Property Let A(ByVal Data As String)
'    mA = Data
'End Property

'Public Property Get A() As String
'    A = mA
'End Property

'Public Property Let A(ByVal Data As String)
'    mA = Data
'End Property

'Public Property Get A() As String
'    A = mA
'End Property

'Public Property Let A(ByVal Data As String)
'    mA = Data
'End Property

'Public Property Get A() As String
'    A = mA
'End Property



