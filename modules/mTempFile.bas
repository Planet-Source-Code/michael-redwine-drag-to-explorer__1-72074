Attribute VB_Name = "mTempFile"
Option Explicit

'**********************************************************************
'* TEMPFILE CREATION STUFF
Public TempFileName As String
Public TempFile As String
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Function CreateTempFile() As String
  'Returns the path and name of the temp file that's created.
  'Generates a temp file name and creates the file for you.
  Dim TempDirPath As String, PathLen As Long

  TempDirPath = Space(255)    ' initialize the buffer to receive the path
  PathLen = GetTempPath(255, TempDirPath)    ' read the path name
  TempDirPath = Left(TempDirPath, PathLen)    ' extract data from the variable
  ' Get a uniquely assigned random file
  CreateTempFile = Space(255)    ' initialize buffer to receive the filename
  Call GetTempFileName(TempDirPath, "TMP", 0, CreateTempFile)    ' get a unique temporary file name
  CreateTempFile = Left(CreateTempFile, InStr(CreateTempFile, vbNullChar) - 1)    ' extract data from the variable
  TempFile = CreateTempFile
  TempFileName = Right(CreateTempFile, Len(CreateTempFile) - InStrRev(CreateTempFile, "\"))
End Function
