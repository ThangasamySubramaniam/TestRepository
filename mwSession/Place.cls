VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Place"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' Place Class - Keeps track of a place in the Enterprise...
' 11 July 2001 ms
'
Option Explicit


Private mPlaceID As String
Private mDriveLetter As String
Private mSubFolderPath As String
Private mRootPath As String
Private mNotation As String


Public Property Get PlaceID() As String
    PlaceID = UCase(mPlaceID)
End Property

Public Property Let PlaceID(ByVal vData As String)
   
    mPlaceID = UCase(vData)
End Property

Public Property Get DriveLetter() As String
    DriveLetter = mDriveLetter
End Property

Public Property Let DriveLetter(ByVal vData As String)
    mDriveLetter = vData
End Property

Public Property Get SubFolderPath() As String
    SubFolderPath = StripFirstLastBackslash(mSubFolderPath)
End Property

Public Property Let SubFolderPath(ByVal vData As String)
    mSubFolderPath = StripFirstLastBackslash(vData)
End Property

Public Property Get Notation() As String
    Notation = mNotation
End Property

Public Property Let Notation(ByVal vData As String)
    mNotation = vData
End Property

Public Property Get RootPath() As String
    RootPath = StripFirstLastBackslash(mRootPath)
End Property

Public Property Let RootPath(ByVal vData As String)
    mRootPath = StripFirstLastBackslash(vData)
End Property




'Public Property Get A() As string
'    A = mA
'End Property

'Public Property Let A(ByVal vData As string)
'    mA = vData
'End Property

'
' Return DriveLetter + SubFolderPath
'
'Public Function GetSubFolderPath() As String
'   On Error GoTo FunctionError
'   ' Test for possible goofy entries in SubFolderPath field...
'   '= c:\foo
'   If InStr(1, ":\", mSubFolderPath) Then
'      GetSubFolderPath = mDriveLetter & ":" & Mid(mSubFolderPath, 3)
'   '= c:foo
'   ElseIf InStr(1, ":", mSubFolderPath) Then
'      GetSubFolderPath = mDriveLetter & ":" & Mid(mSubFolderPath, 2)
'   '= \foo
'   ElseIf Mid(mSubFolderPath, 1, 1) = "\" Then
'      GetSubFolderPath = mDriveLetter & ":" & mSubFolderPath
'   Else
'      '= foo\bla
'      GetSubFolderPath = mDriveLetter & ":\" & mSubFolderPath
'   End If
'   Exit Function
'FunctionError:
'   GetSubFolderPath = ""
'End Function

' Return DriveLetter + SubFolderPath
'
Public Function GetSubFolderPath() As String
   GetSubFolderPath = mDriveLetter & ":" & "\" & RootPath & "\" & SubFolderPath
End Function


Public Function GetFullFilename(FileName As String) As String
   GetFullFilename = mDriveLetter & ":" & "\" & RootPath & "\" & SubFolderPath & "\" & FileName
End Function



Public Function GetSubstitutePath(Optional NewDriveLetter As Variant, Optional NewRootPath _
  As Variant, Optional NewSubFolderPath As Variant) As String
   Dim StrLetter As String
   Dim StrRoot As String
   Dim strPath As String
   
   If IsMissing(NewDriveLetter) Then
      StrLetter = mDriveLetter
   Else
      StrLetter = mID(NewDriveLetter, 1, 1)
   End If
   '
   If IsMissing(NewRootPath) Then
      StrRoot = NewRootPath
   Else
      StrRoot = RootPath
   End If
   '
   If IsMissing(NewSubFolderPath) Then
      strPath = SubFolderPath
   Else
      strPath = NewSubFolderPath
   End If
   
   GetSubstitutePath = StrLetter & ":" & StrRoot & "\" & strPath
  
End Function

'
' Merge DriveLetter + RootPath + SubFolderPath + Filename, return as string
'
'Public Function GetFullFilename(FileName As String) As String
'   On Error GoTo FunctionError
'   ' Test for possible goofy entries in SubFolderPath field...
'   '= c:\foo
'   If InStr(1, ":\", mSubFolderPath) Then
'      GetFullFilename = mDriveLetter & ":" & Mid(mSubFolderPath, 3)
'   '= c:foo
'   ElseIf InStr(1, ":", mSubFolderPath) Then
'      GetFullFilename = mDriveLetter & ":" & Mid(mSubFolderPath, 2)
'   '= \foo
'   ElseIf Mid(mSubFolderPath, 1, 1) = "\" Then
'      GetFullFilename = mDriveLetter & ":" & mSubFolderPath
'   Else
'      '= foo\bla
'      GetFullFilename = mDriveLetter & ":\" & mSubFolderPath
'   End If
'   ' should add check for last slash...
'   GetFullFilename = GetFullFilename & "\" & FileName
'   Exit Function
'FunctionError:
'   GetFullFilename = ""
'End Function

'
' Pass in



Private Function StripFirstLastBackslash(ThePath As String) As String
   Dim i As Integer
   Dim strPath As String
   i = Len(ThePath)
   If i = 0 Then Exit Function
   If mID(ThePath, i, 1) = "\" Then
      strPath = mID(ThePath, 1, i - 1)
   Else
      strPath = ThePath
   End If
   If mID(strPath, 1, 1) = "\" Then
      strPath = mID(strPath, 2)
   End If
   StripFirstLastBackslash = strPath
End Function




