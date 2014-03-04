VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwDotNetForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"WFD_Menu"
Attribute VB_Ext_KEY = "Member0" ,"WFD_Menu"


Option Explicit

'local variable to hold collection
Private mCol As Collection
Private mbIsAdded As Boolean
Private mw_rErrorMsg As String

Public Function Add(FormHandle As Object, VB6FormKey As Long) As mwDotNetForm
   Dim objNewMember As mwDotNetForm
   Dim strFormHandle As String
   
   On Error GoTo add_error
   mw_rErrorMsg = ""
   
   Set objNewMember = New mwDotNetForm
   
   objNewMember.FormHandle = FormHandle
   objNewMember.VB6FormKey = VB6FormKey
      
   strFormHandle = MakeFormHandleKey(VB6FormKey)
   
   mCol.Add objNewMember, strFormHandle
   
   Set Add = objNewMember
   Set objNewMember = Nothing
   mbIsAdded = True
   Exit Function
add_error:
   If Err.Number = 457 Then
      goSession.RaiseError "Error in mwDotNetForms.Add, " & VB6FormKey & " is a duplicate Form entry.", Err.Number, Err.Description
   Else
      goSession.RaiseError "General Error in mwDotNetForms.Add, duplicate entry in mwDotNetForms.", Err.Number, Err.Description
   End If
   If Not objNewMember Is Nothing Then Set objNewMember = Nothing
   mbIsAdded = False
End Function

Public Property Get MakeFormHandleKey(FormHandle As Long) As String
   MakeFormHandleKey = CStr(FormHandle)
End Property

Public Property Get Item(vntIndexKey As Variant) As mwDotNetForm
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property


Public Property Get Count() As Long
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)

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

Public Property Get isAdded() As Boolean
   isAdded = mbIsAdded
End Property

Public Property Get ErrorMsg() As String
   ErrorMsg = mw_rErrorMsg
End Property
