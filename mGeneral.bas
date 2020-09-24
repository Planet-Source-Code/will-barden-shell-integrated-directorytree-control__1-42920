Attribute VB_Name = "mGeneral"
Option Explicit
'
' ----------------------------------------------------------------------------------- '
' File........: mGeneral.bas
' Version.....: 1.1.4
' Author......: Will Barden
' Created.....: 08/01/03
' Modified....: 09/01/03
'
' A module to accompany the Explorer controls. This module is necessary
' because we're subclassing, and that must be done from within a general
' module. Since this module may service more than one Explorer control, the
' easiest way to store all the various properties (old function address,
' object pointer to callback the control, and the shell notification handle)
' is via the SetProp and GetProp APIs.
' Lastly, this module contains general helper functions/subs, plus any
' declarations.
' ----------------------------------------------------------------------------------- '
'
' ----------------------------------------------------------------------------------- '
' Enumerations
' ----------------------------------------------------------------------------------- '
Public Enum FileErrors
   Blank = 0
End Enum
'
Public Enum DirectoryErrors
   DERR_BASE_ERROR = vbObjectError
   DERR_ALREADY_INITIALIZED = DERR_BASE_ERROR + 1
   DERR_SUBCLASS_FAILED = DERR_BASE_ERROR + 2
   DERR_REGISTER_FAILED = DERR_BASE_ERROR + 3
   DERR_NO_DRIVE_LIST = DERR_BASE_ERROR + 4
   DERR_NO_DIRECTORY_LIST = DERR_BASE_ERROR + 5
   DERR_INVALID_DRIVE = DERR_BASE_ERROR + 6
   DERR_INVALID_FOLDER = DERR_BASE_ERROR + 7
End Enum
'
Public Enum DriveTypes
   DRIVE_UNKNOWN = 0
   DRIVE_NO_ROOT_DIR = 1   ' Bad path specified
   DRIVE_REMOVABLE = 2     ' Floppy
   DRIVE_FIXED = 3         ' Hard disk
   DRIVE_REMOTE = 4        ' Network drive
   DRIVE_CDROM = 5         ' CD/DVD
   DRIVE_RAMDISK = 6
End Enum
'
Public Enum BorderTypes
   BDR_NONE = 0
   BDR_RAISED_THIN = &H4
   BDR_RAISED_THICK = &H4 Or &H1
   BDR_SUNKEN_THIN = &H2
   BDR_SUNKEN_THICK = &H2 Or &H8
End Enum
'
' ----------------------------------------------------------------------------------- '
' Consants
' ----------------------------------------------------------------------------------- '
Public Const CF_HDROP                 As Long = 15
' Preferred drop effect: 52134  ' 5=copy, 2=cut
' Filename:              50773  ' for single files?
' Directories are just a single filename.
Public Const CFSTR_PREFERREDDROPEFFECT As String = "Preferred DropEffect"
'
Public Const BF_BOTTOM                As Long = &H8   ' for drawing borders
Public Const BF_LEFT                  As Long = &H1
Public Const BF_RIGHT                 As Long = &H4
Public Const BF_TOP                   As Long = &H2
Public Const BF_RECT                  As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
'
Public Const BUFFER_LEN               As Long = 1024  ' standard string buffer
'
Public Const CSIDL_DESKTOP            As Long = &H0   ' to locate the desktop folder
'
Public Const FIND_ALL_FILES           As String = "*" ' wildcard
'
Public Const FILE_ATTRIBUTE_ARCHIVE   As Long = &H20  ' To retrieve attributes from files and directories
Public Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Public Const FILE_ATTRIBUTE_HIDDEN    As Long = &H2
Public Const FILE_ATTRIBUTE_NORMAL    As Long = &H80
Public Const FILE_ATTRIBUTE_READONLY  As Long = &H1
Public Const FILE_ATTRIBUTE_SYSTEM    As Long = &H4
'
Public Const FLDR_FILE                As String = " "
Public Const FLDR_NORMAL              As String = "F"   ' to avoid key clashes
Public Const FLDR_VIRTUAL             As String = "V"
Public Const FLDR_DRIVE               As String = "D"
Public Const FLDR_INVALID_CHARS       As String = "\/:*?<>|" ' when renaming folders
'
Public Const FO_DELETE                As Long = &H3   ' used when moving files to
Public Const FOF_ALLOWUNDO            As Long = &H40  ' the recycle bin
'
Public Const GMEM_FIXED               As Long = &H0
Public Const GMEM_MOVEABLE            As Long = &H2
'
Public Const GWL_WNDPROC              As Long = (-4)
'
Public Const INVALID_HANDLE_VALUE     As Long = -1    ' Invalid file search handle
'
Public Const MAX_PATH                 As Long = 260
'
Public Const PROP_OLDWINPROC          As String = "OldWinProc"
Public Const PROP_OBJCALLBACK         As String = "ObjCallback"
Public Const PROP_NOTIFY_HANDLE       As String = "ShellNotifyHandle"
Public Const PROP_DESKTOP_PIDL        As String = "DesktopPidl"
'
Public Const SHCNE_MKDIR              As Long = &H8  ' events to receive from the shell
Public Const SHCNE_RMDIR              As Long = &H10
Public Const SHCNE_RENAMEFOLDER       As Long = &H20000
Public Const SHCNE_FOLDER_EVENTS      As Long = (SHCNE_MKDIR) Or (SHCNE_RMDIR) Or (SHCNE_RENAMEFOLDER)
'
Public Const SHCNF_INTERRUPTS         As Long = &H1
Public Const SHCNF_NON_INTERRUPTS     As Long = &H2
'
Public Const WM_USER                  As Long = &H400
Public Const WM_SHELL_NOTIFY          As Long = WM_USER + 1   ' custom message for the shell to send us
'
' ----------------------------------------------------------------------------------- '
' Structs
' ----------------------------------------------------------------------------------- '
Public Type POINTAPI
   x                 As Long
   y                 As Long
End Type
'
Public Type RECT
   Left              As Long
   Top               As Long
   Right             As Long
   Bottom            As Long
End Type
'
Public Type DROPFILES
   pFiles            As Long
   pt                As POINTAPI
   fNC               As Long
   fWide             As Long
End Type
'
Public Type FILETIME
   dwLowDateTime     As Long
   dwHighDateTime    As Long
End Type
'
Public Type SYSTEMTIME
   wYear             As Integer
   wMonth            As Integer
   wDayOfWeek        As Integer
   wDay              As Integer
   wHour             As Integer
   wMinute           As Integer
   wSecond           As Integer
   wMilliseconds     As Integer
End Type
'
Public Type WIN32_FINDDATA
   dwFileAttributes  As Long
   ftCreationTime    As FILETIME
   ftLastAccessTime  As FILETIME
   ftLastWriteTime   As FILETIME
   nFileSizeHigh     As Long
   nFileSizeLow      As Long
   dwReserved0       As Long
   dwReserved1       As Long
   cFileName         As String * MAX_PATH
   cAlternate        As String * 14
End Type
'
Public Type SHFILEOP_STRUCT
   hWnd              As Long
   wFunc             As Long
   pFrom             As String
   pTo               As String
   fFlags            As Integer
   fAborted          As Boolean
   hNameMaps         As Long
   sProgress         As String
End Type
'
Public Type SHNOTIFY_EVENT
   dwItem1           As Long  ' event specific.
   dwItem2           As Long
End Type
'
Public Type SHNOTIFY_REGISTER
   pIdl              As Long  ' pidl of folder to watch.
   fRecursive        As Long  ' flag to indicate watching of subfolders.
End Type
'
' ----------------------------------------------------------------------------------- '
' Apis
' ----------------------------------------------------------------------------------- '
'
' memory handling
Public Declare Sub ApiCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function ApiGlobalAlloc Lib "kernel32.dll" Alias "GlobalAlloc" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function ApiGlobalFree Lib "kernel32.dll" Alias "GlobalFree" (ByVal hMem As Long) As Long
Public Declare Function ApiGlobalLock Lib "kernel32.dll" Alias "GlobalLock" (ByVal hMem As Long) As Long
Public Declare Function ApiGlobalUnlock Lib "kernel32.dll" Alias "GlobalUnlock" (ByVal hMem As Long) As Long
Public Declare Sub ApiCoTaskMemFree Lib "ole32" Alias "CoTaskMemFree" (ByVal pv As Long)
'
' drawing
Public Declare Function ApiDrawEdge Lib "user32.dll" Alias "DrawEdge" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
'
' file searching
Public Declare Function ApiFindClose Lib "kernel32.dll" Alias "FindClose" (ByVal hFindFile As Long) As Long
Public Declare Function ApiFindFirstFile Lib "kernel32.dll" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FINDDATA) As Long
Public Declare Function ApiFindNextFile Lib "kernel32.dll" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FINDDATA) As Long
'
' disk drives
Public Declare Function ApiGetDriveType Lib "kernel32.dll" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function ApiGetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Public Declare Function ApiGetLogicalDriveStrings Lib "kernel32.dll" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'
' drag'n'drop
Public Declare Function ApiDragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
'
' subclassing
Public Declare Function ApiSetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function ApiCallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'
' window properties
Public Declare Function ApiGetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function ApiSetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function ApiRemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
'
' shell functions
Public Declare Function ApiSHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOP_STRUCT) As Long
Public Declare Function ApiSHGetSpecialFolderLocation Lib "shell32" Alias "SHGetSpecialFolderLocation" (ByVal hwndOwner As Long, ByVal nFolder As Long, pIdl As Long) As Long
Public Declare Function ApiSHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pIdl As Long, ByVal pszPath As String) As Long
Public Declare Function ApiSHChangeNotifyRegister Lib "shell32.dll" Alias "#2" (ByVal hWnd As Long, ByVal dwFlags As Long, ByVal wEventsMask As Long, ByVal wMsg As Long, ByVal cItems As Long, lpItems As SHNOTIFY_REGISTER) As Long
Public Declare Function ApiSHChangeNotifyDeregister Lib "shell32.dll" Alias "#4" (ByVal ulID As Long) As Long
Public Declare Sub ApiSHChangeNotify Lib "shell32.dll" Alias "SHChangeNotify" (ByVal wEventId As Long, ByVal uFlags As Long, ByVal dwItem1 As Long, ByVal dwItem2 As Long)
'
Public Declare Function ApiFileTimeToSystemTime Lib "kernel32" Alias "FileTimeToSystemTime" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Public Declare Function ApiFileTimeToLocalFileTime Lib "kernel32" Alias "FileTimeToLocalFileTime" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
'
' ----------------------------------------------------------------------------------- '
' Subclassing methods
' ----------------------------------------------------------------------------------- '
Public Function Subclass(ByVal hWnd As Long, _
                         ByVal lpCallback As Long) As Boolean
Dim lpOldProc As Long
   '
   ' Firstly, check if this window is already subclassed. We don't want to
   ' do it twice..
   If (ApiGetProp(hWnd, PROP_OLDWINPROC) = 0) Then
      '
      ' Now subclass the window.
      lpOldProc = ApiSetWindowLong(hWnd, GWL_WNDPROC, AddressOf SubclassProc)
      '
      ' And check the success, and then save our extra window properties.
      If lpOldProc Then
         '
         Call ApiSetProp(hWnd, PROP_OLDWINPROC, lpOldProc)
         Call ApiSetProp(hWnd, PROP_OBJCALLBACK, lpCallback)
         '
      End If
      '
   End If
   '
   ' Return the success to the caller, so they can react to a failure.
   If lpOldProc Then
      Subclass = True
   End If
   '
End Function
'
Public Sub UnSubclass(ByVal hWnd As Long)
Dim lpOldProc As Long
   '
   ' Retrieve the old window procedure address from the window properties.
   lpOldProc = ApiGetProp(hWnd, PROP_OLDWINPROC)
   If lpOldProc <> 0 Then
      '
      ' Reset the old window proc in place.
      Call ApiSetWindowLong(hWnd, GWL_WNDPROC, lpOldProc)
      '
      ' Remove the extra properties.
      Call ApiRemoveProp(hWnd, PROP_OLDWINPROC)
      Call ApiRemoveProp(hWnd, PROP_OBJCALLBACK)
      '
   End If
   '
End Sub
'
Public Function SubclassProc(ByVal hWnd As Long, _
                             ByVal uMsg As Long, _
                             ByVal wParam As Long, _
                             ByVal lParam As Long) As Long
Dim lpOldProc     As Long
Dim lpCallback    As Long
Dim oCallback     As ucDirectoryTree
   '
   ' Grab the callback object pointer from the window handle, then
   ' raise the event to it. Only do this is we've received a shell change
   ' notification though.
   If uMsg = WM_SHELL_NOTIFY Then
      '
      lpCallback = ApiGetProp(hWnd, PROP_OBJCALLBACK)
      If lpCallback <> 0 Then
         '
         Call ApiCopyMemory(oCallback, lpCallback, 4&)
         Call oCallback.ShellNotifyMsg(wParam, lParam)
         Call ApiCopyMemory(oCallback, 0&, 4&)
         '
      End If
      '
   End If
   '
   ' Call the default window procedure.
   lpOldProc = ApiGetProp(hWnd, PROP_OLDWINPROC)
   If lpOldProc Then
      SubclassProc = ApiCallWindowProc(lpOldProc, hWnd, uMsg, wParam, lParam)
   End If
   '
End Function
'
' ----------------------------------------------------------------------------------- '
' Shell Notify methods
' ----------------------------------------------------------------------------------- '
Public Function RegisterNotify(ByVal hWnd As Long) As Boolean
Dim pIdl   As Long
Dim uNotify As SHNOTIFY_REGISTER
Dim hNotify As Long
   '
   ' Check that this window hasn't been registered already.
   If (ApiGetProp(hWnd, PROP_NOTIFY_HANDLE) = 0) Then
      '
      ' Get a PIDL to point to the Desktop, and setup our request -
      ' we want to watch the entire system for changes.
      Call ApiSHGetSpecialFolderLocation(0&, CSIDL_DESKTOP, pIdl)
      If (pIdl <> 0) Then
         '
         With uNotify
            .pIdl = pIdl
            .fRecursive = True
         End With
         hNotify = ApiSHChangeNotifyRegister(hWnd, _
                                             SHCNF_INTERRUPTS Or SHCNF_NON_INTERRUPTS, _
                                             SHCNE_FOLDER_EVENTS, _
                                             WM_SHELL_NOTIFY, _
                                             1, _
                                             uNotify)
         If hNotify Then
            '
            ' Store the registered handle with the window, since we'll need
            ' it again to unregister notifications.
            Call ApiSetProp(hWnd, PROP_NOTIFY_HANDLE, hNotify)
            Call ApiSetProp(hWnd, PROP_DESKTOP_PIDL, pIdl)
            '
         End If
         '
      Else
         '
         ' For some reason we couldn't get a PIDL to the desktop, so free
         ' up any memory allocated.
         Call ApiCoTaskMemFree(pIdl)
         '
      End If
      '
   End If
   '
   ' Return based on the value of hNotify - it's a handle, we succeeded.
   RegisterNotify = hNotify
   '
End Function
'
Public Sub UnregisterNotify(ByVal hWnd As Long)
Dim hNotify    As Long
Dim pIdl      As Long
   '
   ' Make sure that the window has a valid notify handle stored.
   hNotify = ApiGetProp(hWnd, PROP_NOTIFY_HANDLE)
   If hNotify Then
      '
      ' Unregister our handle.
      Call ApiSHChangeNotifyDeregister(hNotify)
      '
      ' Get the PIDL, and free up any memory used by it.
      pIdl = ApiGetProp(hWnd, PROP_DESKTOP_PIDL)
      If pIdl Then Call ApiCoTaskMemFree(pIdl)
      '
      ' Remove our properties, ready for next time (if there is a next time).
      Call ApiRemoveProp(hWnd, PROP_NOTIFY_HANDLE)
      Call ApiRemoveProp(hWnd, PROP_DESKTOP_PIDL)
      '
   End If
   '
End Sub
'
' ----------------------------------------------------------------------------------- '
' General module functions
' ----------------------------------------------------------------------------------- '
Public Function GetPathFromPidl(ByVal pIdl As Long) As String
Dim sTmp As String
   '
   ' Take a pidl, and fully resolve it to return the string.
   sTmp = String$(BUFFER_LEN, 0)
   Call ApiSHGetPathFromIDList(pIdl, sTmp)
   '
   ' Make sure it's got a slash on the end of it.
   sTmp = AddSlash(Mid$(sTmp, 1, InStr(1, sTmp, Chr$(0)) - 1))
   GetPathFromPidl = sTmp
   '
End Function
'
Public Sub TrimNulls(ByRef sText As String)
Dim lPos As Long
   '
   ' Locate the first instance of a Chr$(0) (NULL) character, and trim
   ' everything after it.
   lPos = InStr(1, sText, Chr$(0))
   If lPos <> 0 Then
      sText = Mid$(sText, 1, lPos - 1)
   End If
   '
End Sub
'
Public Function AddSlash(ByVal sText As String) As String
   '
   ' Return the string with a single trailing slash on the end.
   If LenB(sText) = 0 Then Exit Function
   '
   If Right$(sText, 1) <> "\" Then
      AddSlash = sText & "\"
   Else
      AddSlash = sText
   End If
   '
End Function
'
