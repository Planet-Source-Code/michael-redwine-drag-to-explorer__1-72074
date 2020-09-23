VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   4440
   ClientLeft      =   1125
   ClientTop       =   420
   ClientWidth     =   4530
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   4530
   Begin MSComctlLib.ListView lvw 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   7435
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDropMode     =   1
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public DroppedFolder As String

Private Const VER_PLATFORM_WIN32_NT As Long = 2

Private Type OSVERSIONINFO
  OSVSize As Long
  dwVerMajor As Long
  dwVerMinor As Long
  dwBuildNumber As Long
  PlatformID As Long
  szCSDVersion As String * 128
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
                                      (lpVersionInformation As Any) As Long


Private Sub Form_Load()
  Dim a As Integer
  For a = 1 To 10
    lvw.ListItems.Add , "Key:" & a, "Text:" & a
  Next a
  lvw.View = lvwList
  lvw.OLEDragMode = ccOLEDragAutomatic

  'don't know if it's just being dumb, but the SHCNE stuff doesn't work for the first few seconds when
  'running in Windows Vista VB6 IDE.  must wait for the SHCNE_UPDATEDIR message to show up.  so i disable the
  'listview until we receive that message to prevent me from freaking out when it doesn't work...
  If IsWinVistaIDE = True Then lvw.Enabled = False

  If SubClass(hwnd) Then
    Call SHNotify_Register(hwnd)
  End If
End Sub

Public Sub NotificationReceipt(wParam As Long, lParam As Long)
  Static LastDroppedFile As String
  Dim shns As SHNOTIFYSTRUCT
  Dim dwItem As Long
  Dim DroppedFile As String
  Dim DroppedFileName As String

  CopyMemory shns, ByVal wParam, Len(shns)
  Select Case lParam
    Case SHCNE_UPDATEDIR  'under Vista VB6 IDE, all the SHCNE stuff seems to not work for a few seconds until this message pops up.
      lvw.Enabled = True
    Case SHCNE_CREATE, SHCNE_RENAMEITEM
      'using Vista, SHCNE_CREATE is what we see.  only uses .dwItem1...
      'using XP, SHCNE_RENAMEITEM is what we see.  the path we're interested resides in .dwItem2...
      dwItem = IIf(lParam = SHCNE_RENAMEITEM, shns.dwItem2, shns.dwItem1)
      'TempFileName --- just the temp file name
      'TempFile --- the full path and name of the temp file in the temp folder
      'DroppedFileName --- just the file name of the dropped file
      'DroppedFile --- the full path and name of the temp file in it's new dropped location
      'GetDisplayNameFromPIDL(shns.dwItem1) --- just the file name
      'GetPathFromPIDL(shns.dwItem1) --- the full path and name of the temp file in it's new dropped location
      DroppedFile = GetPathFromPIDL(dwItem)
      DroppedFileName = GetDisplayNameFromPIDL(dwItem)
      Debug.Print "DoppedFile = " & DroppedFile & vbCrLf & vbCrLf & "DroppedFileName = " & DroppedFileName
      'Exit Sub
      If shns.dwItem1 And _
         DroppedFileName = TempFileName And _
         TempFile <> DroppedFile Then
        'our temp file was moved from it's temp folder.  this is the meat and potatoes...
        If LastDroppedFile = DroppedFile Then Exit Sub
        LastDroppedFile = DroppedFile
        DroppedFolder = PathFromFile(DroppedFile)
        Call RecycleFile(DroppedFolder & DroppedFileName)
      End If
  End Select
End Sub

Public Function PathFromFile(ByVal this As String) As String
  PathFromFile = Left(this, InStrRev(this, "\"))
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Call SHNotify_Unregister
  Call UnSubClass(hwnd)
End Sub

Private Function IsFile(ByVal Path As String) As Boolean
  On Error GoTo hell:
  Dim fso
  If Right(Path, 1) = "\" Then Path = Left(Path, Len(Path) - 1)
  Set fso = CreateObject("Scripting.FileSystemObject")
  IsFile = fso.GetFile(Path) <> ""
hell:
  Set fso = Nothing
End Function

Private Sub lvw_OLECompleteDrag(Effect As Long)
  If DroppedFolder = "" Then Exit Sub
  Call MsgBox("The target folder you dragged your listitem to is:" & vbCrLf & vbCrLf & DroppedFolder & vbCrLf & vbCrLf & "From here, you can do whatever you want with it.  For example, if you write an FTP client, you can transfer a file to this folder.  This is just one useful example...", vbInformation, "DroppedFolder")
  'cleanup
  DroppedFolder = ""
  'DroppedFile = ""
  'DroppedFileName = ""
  TempFile = ""
  TempFileName = ""
End Sub

Private Sub lvw_OLESetData(Data As MSComctlLib.DataObject, DataFormat As Integer)
  Data.Files.Add CreateTempFile
End Sub

Private Sub lvw_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
  DroppedFolder = ""
  AllowedEffects = vbDropEffectMove Or vbDropEffectCopy
  Data.SetData , 15
End Sub

Private Function IsWinVistaIDE() As Boolean
  On Error Resume Next
  'returns True if running Windows Vista
  Dim osv As OSVERSIONINFO

  osv.OSVSize = Len(osv)
  If GetVersionEx(osv) = 1 Then
    If (osv.PlatformID = VER_PLATFORM_WIN32_NT) And (osv.dwVerMajor = 6) Then
      Debug.Print 1 / 0
      If Err.Description <> "" Then
        IsWinVistaIDE = True
      End If
    End If
  End If
End Function

