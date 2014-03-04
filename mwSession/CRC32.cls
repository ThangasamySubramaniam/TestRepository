VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRC32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Cyclic Redundancy Check for generating a 32 bit checksum from the contents of a file.

' If you generate a Checksum for a file and send that checksum and a copy of the file to a remote site
' you can then use the original checksum value to determine if the file that arrived on the remote site
' is the same as the file that was sent.

' This code is taken from the VB.NET CRC32 algorithm (then modified a bit by T. McSally)
' provided by Paul (wpsjr1@succeed.net) - Excellent work!

Private crc32Table() As Long
Private Const BUFFER_SIZE As Long = 8192

Private m_sFile As String
Private m_iFile As Integer
Private m_iLen As Long
Private m_iOffset As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Public Function GetFileCheckSum(InFile As String) As Long

   Dim crc32Result As Long
   Dim buffer(0 To BUFFER_SIZE - 1) As Byte
   Dim readSize As Long
   Dim count As Integer
   Dim i As Integer
   Dim iLookup As Integer
   Dim tot As Integer
   
   On Error GoTo ErrorHandler
   
   crc32Result = &HFFFFFFFF
   readSize = BUFFER_SIZE
   
   If OpenFile(InFile) = True Then
      count = ReadChunk(buffer, readSize)
      
      Do While (count > 0)
         For i = 0 To count - 1
            iLookup = (crc32Result And &HFF) Xor buffer(i)
            crc32Result = ((crc32Result And &HFFFFFF00) \ &H100) And 16777215 ' nasty shr 8 with vb :/
            crc32Result = crc32Result Xor crc32Table(iLookup)
         Next i
         count = ReadChunk(buffer, readSize)
      Loop
   End If
   
   GetFileCheckSum = Not (crc32Result)

   Exit Function
   
ErrorHandler:

End Function

Public Function GetByteArrayCrc32(ByRef buffer() As Byte) As Long
   
   Dim crc32Result As Long
   crc32Result = &HFFFFFFFF
      
   Dim i As Integer
   Dim iLookup As Integer
   
   For i = LBound(buffer) To UBound(buffer)
      iLookup = (crc32Result And &HFF) Xor buffer(i)
      crc32Result = ((crc32Result And &HFFFFFF00) \ &H100) And 16777215 ' nasty shr 8 with vb :/
      crc32Result = crc32Result Xor crc32Table(iLookup)
   Next i
   
   GetByteArrayCrc32 = Not (crc32Result)

End Function

Private Sub Class_Initialize()

    ' This is the official polynomial used by CRC32 in PKZip.
    ' Often the polynomial is shown reversed (04C11DB7).
    Dim dwPolynomial As Long
    dwPolynomial = &HEDB88320
    Dim i As Integer, j As Integer

    ReDim crc32Table(256)
    Dim dwCrc As Long

    For i = 0 To 255
        dwCrc = i
        For j = 8 To 1 Step -1
            If (dwCrc And 1) Then
                dwCrc = ((dwCrc And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                dwCrc = dwCrc Xor dwPolynomial
            Else
                dwCrc = ((dwCrc And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
            End If
        Next j
        crc32Table(i) = dwCrc
    Next i

End Sub

Private Sub Class_Terminate()
   Dispose
End Sub

Private Function OpenFile(ByVal sFile As String) As Boolean
   Dim lErr As Long
   
   Dispose
   m_sFile = sFile
   If (FileExists(m_sFile, lErr)) Then
      m_iFile = FreeFile
      Open m_sFile For Binary Access Read Lock Write As #m_iFile
      m_iLen = LOF(m_iFile)
      OpenFile = True
   Else
      OpenFile = False
   End If
End Function

Private Function FileExists(ByVal sFile As String, ByRef lErr As Long) As Boolean
   
   lErr = 0
   On Error Resume Next
   Dim sDir As String
   sDir = Dir(sFile)
   lErr = err.Number
   On Error GoTo 0
   
   If (lErr = 0) Then
      If (Len(sDir) > 0) Then
         FileExists = True
      Else
         lErr = 53
      End If
   End If
   
End Function

Private Function ReadChunk(buffer() As Byte, ByVal readSize As Long) As Long
   
   Dim lReadSize As Long
   lReadSize = readSize
   If (m_iOffset + lReadSize >= m_iLen) Then
      readSize = m_iLen - m_iOffset
      If (readSize > 0) Then
         ReDim newBuffer(0 To readSize - 1) As Byte
         Get #m_iFile, , newBuffer
         CopyMemory buffer(0), newBuffer(0), readSize
      Else
         Dispose
      End If
      m_iOffset = m_iOffset + readSize
   Else
      ' Can read
      Get #m_iFile, , buffer
      m_iOffset = m_iOffset + readSize
   End If
   ReadChunk = readSize
   
End Function

Private Sub Dispose()
   If (m_iFile) Then
      Close #m_iFile
      m_iFile = 0
   End If
End Sub



