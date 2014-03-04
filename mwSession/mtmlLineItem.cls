VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mtmlLineItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' mtmlLineItem - MTML Requisition Line Item
' 23 March 2002 ms
'

Option Explicit

Private mLineItem As String
Private mDescription As String
Private mMfgPartNo As String
Private mDrawingNumber As String
Private mQuantity As Double
Private mMonetaryAmount As Double
Private mMeasureUnitQualifier As String
Private mComment As String
'Optional Section Info
Private mEqptMfg As String
Private mEqptModel As String
Private mEqptDescription As String
Private mEqptSerialNumber As String

Private moSectionManufacturer As mtmlSectionManufacturer
Private moTypeCodes


Public Property Let LineItem(ByVal Data As String)
    mLineItem = Data
End Property

Public Property Get LineItem() As String
    LineItem = mLineItem
End Property

Public Property Let Description(ByVal Data As String)
    mDescription = Data
End Property

Public Property Get Description() As String
    Description = mDescription
End Property

Public Property Let Quantity(ByVal Data As Double)
    mQuantity = Data
End Property

Public Property Get Quantity() As Double
    Quantity = mQuantity
End Property

Public Property Let MonetaryAmount(ByVal Data As String)
    mMonetaryAmount = Data
End Property

Public Property Get MonetaryAmount() As String
    MonetaryAmount = mMonetaryAmount
End Property

Public Property Let MeasureUnitQualifier(ByVal Data As String)
    mMeasureUnitQualifier = Data
End Property

Public Property Get MeasureUnitQualifier() As String
    MeasureUnitQualifier = mMeasureUnitQualifier
End Property

Public Property Let MfgPartNo(ByVal Data As String)
    mMfgPartNo = Data
End Property

Public Property Get MfgPartNo() As String
    MfgPartNo = mMfgPartNo
End Property

Public Property Let DrawingNumber(ByVal Data As String)
    mDrawingNumber = Data
End Property

Public Property Get DrawingNumber() As String
    DrawingNumber = mDrawingNumber
End Property


Public Property Let Comment(ByVal Data As String)
    mComment = Data
End Property

Public Property Get Comment() As String
    Comment = mComment
End Property

Public Property Let EqptMfg(ByVal Data As String)
    mEqptMfg = Data
End Property

Public Property Get EqptMfg() As String
    EqptMfg = mEqptMfg
End Property

Public Property Let EqptDescription(ByVal Data As String)
    mEqptDescription = Data
End Property

Public Property Get EqptDescription() As String
    EqptDescription = mEqptDescription
End Property

Public Property Let EqptSerialNumber(ByVal Data As String)
    mEqptSerialNumber = Data
End Property

Public Property Get EqptSerialNumber() As String
    EqptSerialNumber = mEqptSerialNumber
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


