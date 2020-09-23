Attribute VB_Name = "mShellFileOps"
Option Explicit

'**********************************************************************
'* SHELL FILE OPERATIONS STUFF, USED FOR MOVING TO RECYCLE BIN
Public Type SHFILEOPSTRUCT
  hwnd As Long
  wFunc As Long
  pFrom As String
  pTo As String
  fFlags As Integer
  fAnyOperationsAborted As Long
  hNameMappings As Long
  lpszProgressTitle As String
End Type    ' SHFILEOPSTRUCT

Public Const FO_MOVE As Long = &H1&
Public Const FO_COPY As Long = &H2&
Public Const FO_DELETE As Long = &H3&
Public Const FO_RENAME As Long = &H4&

Public Const FOF_CREATEPROGRESSDLG As Integer = &H0
Public Const FOF_MULTIDESTFILES = &H1
Public Const FOF_CONFIRMMOUSE As Integer = &H2
Public Const FOF_SILENT As Integer = &H4
Public Const FOF_RENAMEONCOLLISION As Integer = &H8
Public Const FOF_NOCONFIRMATION As Integer = &H10
Public Const FOF_WANTMAPPINGHANDLE As Integer = &H20
Public Const FOF_ALLOWUNDO As Integer = &H40
Public Const FOF_FILESONLY As Integer = &H80
Public Const FOF_SIMPLEPROGRESS As Integer = &H100
Public Const FOF_NOCONFIRMMKDIR As Integer = &H200
Public Const FOF_NOERRORUI As Integer = &H400
Public Const FOF_NOCOPYSECURITYATTRIBS As Integer = &H800

Public Declare Function SHFileOperation& Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As Any)

Public Function ShellCopy(ByVal src As String, ByVal dest As String)
  Dim fos As SHFILEOPSTRUCT    ' structure to pass to the function
  Dim sa(1 To 32) As Byte    ' byte array to make structure properly sized
  Dim tmp() As String    ' return value

  With fos
    .hwnd = frmMain.hwnd
    .wFunc = FO_COPY
    .pFrom = src
    .pTo = dest

    .fFlags = FOF_FILESONLY
    'determine if there are multiple files to copy.  if so, add the following flag.
    If UBound(Split(src, vbNullChar)) >= 3 Then .fFlags = .fFlags Or FOF_MULTIDESTFILES
  End With

  CopyMemory sa(1), fos, LenB(fos)
  CopyMemory sa(19), sa(21), 12

  'Call BringWindowToTop(frmMain.hWnd)
  ShellCopy = SHFileOperation(sa(1)) = 0

  CopyMemory sa(21), sa(19), 12
  CopyMemory fos, sa(1), Len(fos)
End Function

Public Sub RecycleFile(ByVal file As String)
  Dim DelFileOp As SHFILEOPSTRUCT

  With DelFileOp
    .hwnd = 0&    'not concerned with this cuz not concerned with callbacks.
    .wFunc = FO_DELETE
    .pFrom = file
    .fFlags = FOF_NOCONFIRMATION Or FOF_FILESONLY Or FOF_SILENT Or FOF_NOERRORUI
  End With
  Call SHFileOperation(DelFileOp)
End Sub


