VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwZipItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"ZipItem"
Attribute VB_Ext_KEY = "Member0" ,"ZipItem"
'local variable to hold collection
Private mCol As Collection
Private moParent As Zipper

Public Property Set Parent(oParent As Zipper)
   If moParent Is Nothing Then
      Set moParent = oParent
   End If
End Property

Public Property Get Parent() As Zipper
   Set Parent = moParent
End Property


Public Function Add(FileName As String, FilesSize As Long, _
  FileDate As Date, HasFilePath As Boolean, sKey As String) As mwZipItem
   'create a new object
   On Error GoTo add_error
   Dim objNewMember As mwZipItem
   
   Set objNewMember = New mwZipItem
   mCol.Add objNewMember, sKey
   
   objNewMember.FileName = FileName
   objNewMember.FileSize = FileSize
   objNewMember.FileDate = FileDate
   objNewMember.HasFilePath = HasFilePath
    
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
    Exit Function
add_error:
    Set objNewMember = Nothing
    If err.Number <> 457 Then
      MsgBox "Error in mwZipItems.Add"
   End If
End Function

Public Property Get Item(vntIndexKey As Variant) As mwZipItem
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

