VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mtmlSectionManufacturer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' mtmlSectionManufacturer -
' 23 March, 2002 ms


Option Explicit

Private mSectionManufacturer As String
Private mName As String
Private mModelNumber As String
Private mSerialNumber As String
Private mDescription As String

Public Property Let SectionManufacturer(ByVal Data As String)
    mSectionManufacturer = Data
End Property

Public Property Get SectionManufacturer() As String
    SectionManufacturer = mSectionManufacturer
End Property

Public Property Let Name(ByVal Data As String)
    mName = Data
End Property

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let ModelNumber(ByVal Data As String)
    mModelNumber = Data
End Property

Public Property Get ModelNumber() As String
    ModelNumber = mModelNumber
End Property

Public Property Let SerialNumber(ByVal Data As String)
    mSerialNumber = Data
End Property

Public Property Get SerialNumber() As String
    SerialNumber = mSerialNumber
End Property

Public Property Let Description(ByVal Data As String)
    mDescription = Data
End Property

Public Property Get Description() As String
    Description = mDescription
End Property

'Public Property Let A(ByVal Data As String)
'    mA = Data
'End Property

'Public Property Get A() As String
'    A = mA
'End Property


