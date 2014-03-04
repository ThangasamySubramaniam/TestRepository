VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwVoyEventTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable to hold collection
Private mCol As Collection

Public Function Add(ID As Long, Optional sKey As String) As mwVoyEventType
    'create a new object
   On Error GoTo FunctionError
   
   Dim objNewMember As mwVoyEventType
   Set objNewMember = New mwVoyEventType
   objNewMember.ID = ID
   mCol.Add objNewMember, str(ID)


   'return the object created
   Set Add = objNewMember
   Set objNewMember = Nothing
   Exit Function
FunctionError:

End Function

Public Property Get Item(vntIndexKey As Variant) As mwVoyEventType
  On Error Resume Next
  Set Item = mCol(vntIndexKey)
  Exit Property
End Property



Public Property Get Count() As Long
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub





