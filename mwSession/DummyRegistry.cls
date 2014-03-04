VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DummyRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"regErrors"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

   Private mBaseRegistry As String


'******* Public Interface *******************************

Public Property Let BaseRegistry(ByVal Data As String)
    mBaseRegistry = Data
End Property

Public Property Get BaseRegistry() As String
    BaseRegistry = mBaseRegistry
End Property



Public Function GetRegInt(Key As String, DefaultValue As Integer) As Integer
   
   On Error GoTo FunctionError
   
   GetRegInt = DefaultValue
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.DummyRegistry.GetRegInt. ", Err.Number, Err.Description
End Function
Public Function GetRegLong(Key As String, DefaultValue As Long) As Long
   
   On Error GoTo FunctionError
   
   GetRegLong = DefaultValue
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.DummyRegistry.GetRegLong. ", Err.Number, Err.Description
End Function
Public Function GetRegStr(Key As String, DefaultValue As String) As String
   
   On Error GoTo FunctionError
   
   GetRegStr = DefaultValue
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.DummyRegistry.GetRegStr. ", Err.Number, Err.Description
End Function

Public Function GetReg(Key As String) As String
   GetReg = ""
End Function

Public Function SetReg(ByVal Key As String, ByVal value As String) As String
   SetReg = value
End Function


' Functions for getting & setting grid column widths & positions.

Public Function ugGetWidth(GridName As String, FieldNum As Integer, DefaultWidth As Integer) As Integer
   On Error GoTo FunctionError
   
   ugGetWidth = DefaultWidth
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.DummyRegistry.ugGetWidth. ", Err.Number, Err.Description

End Function

Public Function ugSetWidth(GridName As String, FieldNum As Integer, Width As Integer) As Integer
   On Error GoTo FunctionError
   
   ugSetWidth = Width
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.DummyRegistry.ugSetWidth. ", Err.Number, Err.Description

End Function

Public Function ugGetPosition(GridName As String, FieldNum As Integer, DefaultPos As Integer) As Integer
   On Error GoTo FunctionError
   
   ugGetPosition = DefaultPos
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.DummyRegistry.ugGetPosition. ", Err.Number, Err.Description

End Function

Public Function ugSetPosition(GridName As String, FieldNum As Integer, Position As Integer) As Integer
   
   On Error GoTo FunctionError
   
   ugSetPosition = Position
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.DummyRegistry.ugSetPosition. ", Err.Number, Err.Description

End Function

