VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwZipItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'mwZipItem - class holds properties about a zip file item

Private mFileName As String
Private mFileSize As Long
Private mFileDate As Date
Private mHasFilePath As Boolean

Public Property Let FileName(ByVal Data As String)
    mFileName = Data
End Property

Public Property Get FileName() As String
    FileName = mFileName
End Property

Public Property Let FileSize(ByVal Data As Long)
    mFileSize = Data
End Property

Public Property Get FileSize() As Long
    FileSize = mFileSize
End Property

Public Property Let FileDate(ByVal Data As Date)
    mFileDate = Data
End Property

Public Property Get FileDate() As Date
    FileDate = mFileDate
End Property

Public Property Let HasFilePath(ByVal Data As Boolean)
    mHasFilePath = Data
End Property

Public Property Get HasFilePath() As Boolean
    HasFilePath = mHasFilePath
End Property



