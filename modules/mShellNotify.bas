Attribute VB_Name = "mShellNotify"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2009 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Brought to you by Brad Martinez
'  http://www.mvps.org/btmtz/
'  http://www.mvps.org/ccrp/
'
'Demonstrates how to receive shell change
'notifications (ala "what happens when the
'SHChangeNotify API is called?")
'
'Interpretation of the shell's undocumented
'functions SHChangeNotifyRegister (ordinal 2)
'and SHChangeNotifyDeregister (ordinal 4) would
'not have been possible without the assistance of
'James Holderness. For a complete (and probably
'more accurate) overview of shell change notifications,
'please refer to James'"Shell Notifications" page at
'http://www.geocities.com/SiliconValley/4942/
'------------------------------------------------------
Public Const MAX_PATH As Long = 260
'Defined as an HRESULT that corresponds to S_OK.
Public Const ERROR_SUCCESS As Long = 0
Public Type SHFILEINFO   'shfi
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type
'If pidl is invalid, SHGetFileInfoPidl can very easily blow up when filling the szDisplayName and szTypeName string members of the SHFILEINFO struct
Public Type SHFILEINFOBYTE   'sfib
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName(1 To MAX_PATH) As Byte
  szTypeName(1 To 80) As Byte
End Type
'Special folder values for
'SHGetSpecialFolderLocation and
'SHGetSpecialFolderPath (Shell32.dll v4.71)
Public Enum SHSpecialFolderIDs
  CSIDL_DESKTOP = &H0
  CSIDL_INTERNET = &H1
  CSIDL_PROGRAMS = &H2
  CSIDL_CONTROLS = &H3
  CSIDL_PRINTERS = &H4
  CSIDL_PERSONAL = &H5
  CSIDL_FAVORITES = &H6
  CSIDL_STARTUP = &H7
  CSIDL_RECENT = &H8
  CSIDL_SENDTO = &H9
  CSIDL_BITBUCKET = &HA
  CSIDL_STARTMENU = &HB
  CSIDL_DESKTOPDIRECTORY = &H10
  CSIDL_DRIVES = &H11
  CSIDL_NETWORK = &H12
  CSIDL_NETHOOD = &H13
  CSIDL_FONTS = &H14
  CSIDL_TEMPLATES = &H15
  CSIDL_COMMON_STARTMENU = &H16
  CSIDL_COMMON_PROGRAMS = &H17
  CSIDL_COMMON_STARTUP = &H18
  CSIDL_COMMON_DESKTOPDIRECTORY = &H19
  CSIDL_APPDATA = &H1A
  CSIDL_PRINTHOOD = &H1B
  CSIDL_ALTSTARTUP = &H1D        ''DBCS
  CSIDL_COMMON_ALTSTARTUP = &H1E  ''DBCS
  CSIDL_COMMON_FAVORITES = &H1F
  CSIDL_INTERNET_CACHE = &H20
  CSIDL_COOKIES = &H21
  CSIDL_HISTORY = &H22
End Enum
Enum SHGFI_FLAGS
  SHGFI_LARGEICON = &H0           'sfi.hIcon is large icon
  SHGFI_SMALLICON = &H1           'sfi.hIcon is small icon
  SHGFI_OPENICON = &H2            'sfi.hIcon is open icon
  SHGFI_SHELLICONSIZE = &H4       'sfi.hIcon is shell size (not system size), rtns BOOL
  SHGFI_PIDL = &H8                'pszPath is pidl, rtns BOOL
  SHGFI_USEFILEATTRIBUTES = &H10  'parent pszPath exists, rtns BOOL
  SHGFI_ICON = &H100              'fills sfi.hIcon, rtns BOOL, use DestroyIcon
  SHGFI_DISPLAYNAME = &H200       'isf.szDisplayName is filled, rtns BOOL
  SHGFI_TYPENAME = &H400          'isf.szTypeName is filled, rtns BOOL
  SHGFI_ATTRIBUTES = &H800        'rtns IShellFolder::GetAttributesOf  SFGAO_* flags
  SHGFI_ICONLOCATION = &H1000     'fills sfi.szDisplayName with filename containing the icon, rtns BOOL
  SHGFI_EXETYPE = &H2000          'rtns two ASCII chars of exe type
  SHGFI_SYSICONINDEX = &H4000     'sfi.iIcon is sys il icon index, rtns hImagelist
  SHGFI_LINKOVERLAY = &H8000&     'add shortcut overlay to sfi.hIcon
  SHGFI_SELECTED = &H10000        'sfi.hIcon is selected icon
End Enum
Public Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
'Frees memory allocated by the shell (pidls)
Public Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
'Retrieves the location of a special (system) folder. Returns ERROR_SUCCESS if successful or an OLE-defined error result otherwise.
Public Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As SHSpecialFolderIDs, pidl As Long) As Long
'Converts an item identifier list to a file system path. Returns TRUE if successful or FALSE if an error occurs, for example, if the location specified by the pidl parameter is not part of the file system.
Public Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
'Retrieves information about an object in the file system, such as a file, a folder, a directory, or a drive root.
Public Declare Function SHGetFileInfoPidl Lib "shell32" Alias "SHGetFileInfoA" (ByVal pidl As Long, ByVal dwFileAttributes As Long, psfib As SHFILEINFOBYTE, ByVal cbFileInfo As Long, ByVal uFlags As SHGFI_FLAGS) As Long
Public Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As SHGFI_FLAGS) As Long
Private Const WM_NCDESTROY As Long = &H82
Private Const GWL_WNDPROC As Long = (-4)
Private Const OLDWNDPROC As String = "OldWndProc"
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'the one and only shell change notification handle for the desktop folder
Private m_hSHNotify As Long
'the desktop's pidl
Private m_pidlDesktop As Long
'User defined notification message sent to the specified window's window proc.
Public Const WM_SHNOTIFY = &H401

'------------------------------------------------------
Public Type PIDLSTRUCT
  'Fully qualified pidl (relative to the desktop folder) of the folder to monitor changes in. 0 can also be specified for the desktop folder.
  pidl As Long
  'Value specifying whether changes in the folder's subfolders trigger a change notification event.
  bWatchSubFolders As Long
End Type
Public Declare Function SHChangeNotifyRegister Lib "shell32" Alias "#2" (ByVal hwnd As Long, ByVal uFlags As SHCN_ItemFlags, ByVal dwEventID As SHCN_EventIDs, ByVal uMsg As Long, ByVal cItems As Long, lpps As PIDLSTRUCT) As Long
'If successful, SHChangeNotifyRegister returns a notification handle which must be passed to SHChangeNotifyDeregister when no longer used. Returns 0 otherwise.
'Once the specified message is registered with SHChangeNotifyRegister, the specified window's function proc will be notified by the shell of the specified event in (and under) the folder(s) specified in a pidl.
'On message receipt, wParam points to a SHNOTIFYSTRUCT and lParam contains the event's ID value.
'The values in dwItem1 and dwItem2 are event specific. See the description of the values for the wEventId parameter of the documented SHChangeNotify API function.
Public Type SHNOTIFYSTRUCT
  dwItem1 As Long
  dwItem2 As Long
End Type
'...?
'Public Declare Function SHChangeNotifyUpdateEntryList Lib "shell32" Alias "#5" (ByVal hNotify As Long, ByVal Unknown As Long, ByVal cItem As Long, lpps As PIDLSTRUCT) As Boolean
'Public Declare Function SHChangeNotifyReceive Lib "shell32" Alias "#5" (ByVal hNotify As Long, ByVal uFlags As SHCN_ItemFlags, ByVal dwItem1 As Long, ByVal dwItem2 As Long) As Long
'Closes the notification handle returned from a call to SHChangeNotifyRegister. Returns True if successful,False otherwise.
Public Declare Function SHChangeNotifyDeregister Lib "shell32" Alias "#4" (ByVal hNotify As Long) As Boolean
'------------------------------------------------------
'This function should be called by any app that changes anything in the shell. The shell will then notify each "notification registered" window of this action.
Public Declare Sub SHChangeNotify Lib "shell32" (ByVal wEventId As SHCN_EventIDs, ByVal uFlags As SHCN_ItemFlags, ByVal dwItem1 As Long, ByVal dwItem2 As Long)
'Shell notification event IDs
Public Enum SHCN_EventIDs
  SHCNE_RENAMEITEM = &H1          '(D) A non-folder item has been renamed.
  SHCNE_CREATE = &H2              '(D) A non-folder item has been created.
  SHCNE_DELETE = &H4              '(D) A non-folder item has been deleted.
  SHCNE_MKDIR = &H8               '(D) A folder item has been created.
  SHCNE_RMDIR = &H10              '(D) A folder item has been removed.
  SHCNE_MEDIAINSERTED = &H20      '(G) Storage media has been inserted into a drive.
  SHCNE_MEDIAREMOVED = &H40       '(G) Storage media has been removed from a drive.
  SHCNE_DRIVEREMOVED = &H80       '(G) A drive has been removed.
  SHCNE_DRIVEADD = &H100          '(G) A drive has been added.
  SHCNE_NETSHARE = &H200          'A folder on the local computer is being shared via the network.
  SHCNE_NETUNSHARE = &H400        'A folder on the local computer is no longer being shared via the network.
  SHCNE_ATTRIBUTES = &H800        '(D) The attributes of an item or folder have changed.
  SHCNE_UPDATEDIR = &H1000        '(D) The contents of an existing folder have changed, but the folder still exists and has not been renamed.
  SHCNE_UPDATEITEM = &H2000       '(D) An existing non-folder item has changed, but the item still exists and has not been renamed.
  SHCNE_SERVERDISCONNECT = &H4000  'The computer has disconnected from a server.
  SHCNE_UPDATEIMAGE = &H8000&     '(G) An image in the system image list has changed.
  SHCNE_DRIVEADDGUI = &H10000     '(G) A drive has been added and the shell should create a new window for the drive.
  SHCNE_RENAMEFOLDER = &H20000    '(D) The name of a folder has changed.
  SHCNE_FREESPACE = &H40000       '(G) The amount of free space on a drive has changed.
  #If (WIN32_IE >= &H400) Then
  SHCNE_EXTENDED_EVENT = &H4000000  '(G) Not currently used.
  #End If
  SHCNE_ASSOCCHANGED = &H8000000   '(G) A file type association has changed.
  SHCNE_DISKEVENTS = &H2381F       '(D) Specifies a combination of all of the disk event identifiers.
  SHCNE_GLOBALEVENTS = &HC0581E0   '(G) Specifies a combination of all of the global event identifiers.
  SHCNE_ALLEVENTS = &H7FFFFFFF
  SHCNE_INTERRUPT = &H80000000     'The specified event occurred as a result of a system
  'interrupt. It is stripped out before the clients of SHCNNotify_ see it.
End Enum

#If (WIN32_IE >= &H400) Then
  Public Const SHCNEE_ORDERCHANGED = &H2  'dwItem2 is the pidl of the changed folder
#End If

'Notification flags uFlags & SHCNF_TYPE is an ID which indicates what dwItem1 and dwItem2 mean
Public Enum SHCN_ItemFlags
  SHCNF_IDLIST = &H0         'LPITEMIDLIST
  SHCNF_PATHA = &H1          'path name
  SHCNF_PRINTERA = &H2       'printer friendly name
  SHCNF_DWORD = &H3          'DWORD
  SHCNF_PATHW = &H5          'path name
  SHCNF_PRINTERW = &H6       'printer friendly name
  SHCNF_TYPE = &HFF
  'Flushes the system event buffer. The function does not return until the system is finished processing the given event.
  SHCNF_FLUSH = &H1000
  'Flushes the system event buffer. The function returns immediately regardless of whether the system is finished processing the given event.
  SHCNF_FLUSHNOWAIT = &H2000
  #If UNICODE Then
  SHCNF_PATH = SHCNF_PATHW
  SHCNF_PRINTER = SHCNF_PRINTERW
  #Else
  SHCNF_PATH = SHCNF_PATHA
  SHCNF_PRINTER = SHCNF_PRINTERA
  #End If
End Enum


Public Function SHNotify_Register(hwnd As Long) As Boolean
  'Registers the one and only shell change notification.
  Dim ps As PIDLSTRUCT

  'If we don't already have a notification going...
  If (m_hSHNotify = 0) Then
    'Get the pidl for the desktop folder.
    m_pidlDesktop = GetPIDLFromFolderID(0, CSIDL_DESKTOP)
    If m_pidlDesktop Then
      'Fill the one and only PIDLSTRUCT, we're watching desktop and all of the its subfolders, everything...
      ps.pidl = m_pidlDesktop
      ps.bWatchSubFolders = True
      'Register the notification, specifying that we want the dwItem1 and dwItem2 members of the SHNOTIFYSTRUCT to be pidls. We're watching all events.
      m_hSHNotify = SHChangeNotifyRegister(hwnd, SHCNF_TYPE Or SHCNF_IDLIST, SHCNE_ALLEVENTS Or SHCNE_INTERRUPT, WM_SHNOTIFY, 1, ps)
      SHNotify_Register = CBool(m_hSHNotify)
    Else
      'If something went wrong...
      Call CoTaskMemFree(m_pidlDesktop)
    End If
  End If
End Function


Public Function SHNotify_Unregister() As Boolean
  'Unregisters the one and only shell change notification.
  'If we have a registered notification handle.
  If m_hSHNotify Then
    'Unregster it. If the call is successful, zero the handle's variable, free and zero the the desktop's pidl.
    If SHChangeNotifyDeregister(m_hSHNotify) Then
      m_hSHNotify = 0
      Call CoTaskMemFree(m_pidlDesktop)
      m_pidlDesktop = 0
      SHNotify_Unregister = True
    End If
  End If
End Function


Public Function SHNotify_GetEventStr(dwEventID As Long) As String
  'Returns the event string associated with the specified event ID value.
  Dim sEvent As String

  Select Case dwEventID
    Case SHCNE_RENAMEITEM: sEvent = "SHCNE_RENAMEITEM"  '&H1
    Case SHCNE_CREATE: sEvent = "SHCNE_CREATE"  '&H2
    Case SHCNE_DELETE: sEvent = "SHCNE_DELETE"  '&H4
    Case SHCNE_MKDIR: sEvent = "SHCNE_MKDIR"  '&H8
    Case SHCNE_RMDIR: sEvent = "SHCNE_RMDIR"  '&H10
    Case SHCNE_MEDIAINSERTED: sEvent = "SHCNE_MEDIAINSERTED"  '&H20
    Case SHCNE_MEDIAREMOVED: sEvent = "SHCNE_MEDIAREMOVED"  '&H40
    Case SHCNE_DRIVEREMOVED: sEvent = "SHCNE_DRIVEREMOVED"  '&H80
    Case SHCNE_DRIVEADD: sEvent = "SHCNE_DRIVEADD"  '&H100
    Case SHCNE_NETSHARE: sEvent = "SHCNE_NETSHARE"  '&H200
    Case SHCNE_NETUNSHARE: sEvent = "SHCNE_NETUNSHARE"  '&H400
    Case SHCNE_ATTRIBUTES: sEvent = "SHCNE_ATTRIBUTES"  '&H800
    Case SHCNE_UPDATEDIR: sEvent = "SHCNE_UPDATEDIR"  '&H1000
    Case SHCNE_UPDATEITEM: sEvent = "SHCNE_UPDATEITEM"  '&H2000
    Case SHCNE_SERVERDISCONNECT: sEvent = "SHCNE_SERVERDISCONNECT"  '&H4000
    Case SHCNE_UPDATEIMAGE: sEvent = "SHCNE_UPDATEIMAGE"  '&H8000&
    Case SHCNE_DRIVEADDGUI: sEvent = "SHCNE_DRIVEADDGUI"  '&H10000
    Case SHCNE_RENAMEFOLDER: sEvent = "SHCNE_RENAMEFOLDER"  '&H20000
    Case SHCNE_FREESPACE: sEvent = "SHCNE_FREESPACE"  '&H40000
      #If (WIN32_IE >= &H400) Then
      Case SHCNE_EXTENDED_EVENT: sEvent = "SHCNE_EXTENDED_EVENT"  '&H4000000
      #End If
    Case SHCNE_ASSOCCHANGED: sEvent = "SHCNE_ASSOCCHANGED"  '&H8000000
    Case SHCNE_DISKEVENTS: sEvent = "SHCNE_DISKEVENTS"  '&H2381F
    Case SHCNE_GLOBALEVENTS: sEvent = "SHCNE_GLOBALEVENTS"  '&HC0581E0
    Case SHCNE_ALLEVENTS: sEvent = "SHCNE_ALLEVENTS"  '&H7FFFFFFF
    Case SHCNE_INTERRUPT: sEvent = "SHCNE_INTERRUPT"  '&H80000000
  End Select
  SHNotify_GetEventStr = sEvent
End Function

Public Function SubClass(hwnd As Long) As Boolean
  Dim lpfnOld As Long, fSuccess As Boolean

  If (GetProp(hwnd, OLDWNDPROC) = 0) Then
    lpfnOld = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
    If lpfnOld Then
      fSuccess = SetProp(hwnd, OLDWNDPROC, lpfnOld)
    End If
  End If
  If fSuccess Then
    SubClass = True
  Else
    If lpfnOld Then Call UnSubClass(hwnd)
    MsgBox "Unable to successfully subclass &H" & Hex(hwnd), vbCritical
  End If
End Function


Public Function UnSubClass(hwnd As Long) As Boolean
  Dim lpfnOld As Long

  lpfnOld = GetProp(hwnd, OLDWNDPROC)
  If lpfnOld Then
    If RemoveProp(hwnd, OLDWNDPROC) Then
      UnSubClass = SetWindowLong(hwnd, GWL_WNDPROC, lpfnOld)
    End If
  End If
End Function

Public Function GetPIDLFromFolderID(hOwner As Long, nFolder As SHSpecialFolderIDs) As Long
  'Returns an absolute pidl (relative to the desktop) from a special folder's ID. (Calling proc is responsible for freeing the pidl)
  'hOwner - handle of window that will own any displayed msg boxes
  'nFolder  - special folder ID
  Dim pidl As Long

  If SHGetSpecialFolderLocation(hOwner, nFolder, pidl) = ERROR_SUCCESS Then
    GetPIDLFromFolderID = pidl
  End If
End Function


Public Function GetDisplayNameFromPIDL(pidl As Long) As String
  'If successful returns the specified absolute pidl's displayname, returns an empty string otherwise.
  Dim sfib As SHFILEINFOBYTE

  If SHGetFileInfoPidl(pidl, 0, sfib, Len(sfib), SHGFI_PIDL Or SHGFI_DISPLAYNAME) Then
    GetDisplayNameFromPIDL = GetStrFromBufferA(StrConv(sfib.szDisplayName, vbUnicode))
  End If
End Function


Public Function GetPathFromPIDL(pidl As Long) As String
  'Returns a path from only an absolute pidl (relative to the desktop).
  Dim sPath As String * MAX_PATH

  'SHGetPathFromIDList rtns TRUE (1), if successful, FALSE (0) if not
  If SHGetPathFromIDList(pidl, sPath) Then
    GetPathFromPIDL = GetStrFromBufferA(sPath)
  End If
End Function


Public Function GetStrFromBufferA(sz As String) As String
  'Return the string before first null char encountered (if any) from an ANSII string. If no null, return the string passed
  If InStr(sz, vbNullChar) Then
    GetStrFromBufferA = Left$(sz, InStr(sz, vbNullChar) - 1)
  Else
    GetStrFromBufferA = sz
  End If
End Function


Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Select Case uMsg
    Case WM_SHNOTIFY
      Call frmMain.NotificationReceipt(wParam, lParam)
    Case WM_NCDESTROY
      Call UnSubClass(hwnd)
      MsgBox "Unsubclassed &H" & Hex(hwnd), vbCritical, "WindowProc Error"
  End Select
  WindowProc = CallWindowProc(GetProp(hwnd, OLDWNDPROC), hwnd, uMsg, wParam, lParam)
End Function

