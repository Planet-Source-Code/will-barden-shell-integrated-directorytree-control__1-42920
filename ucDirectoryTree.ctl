VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ucDirectoryTree 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ucDirectoryTree.ctx":0000
   Begin MSComctlLib.ImageList DirectoryImages 
      Left            =   3240
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucDirectoryTree.ctx":0312
            Key             =   "Fixed"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucDirectoryTree.ctx":03BE
            Key             =   "Removable"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucDirectoryTree.ctx":0482
            Key             =   "CdRom"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucDirectoryTree.ctx":05A6
            Key             =   "Remote"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucDirectoryTree.ctx":0676
            Key             =   "RamDisk"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucDirectoryTree.ctx":0786
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucDirectoryTree.ctx":0B32
            Key             =   "FolderCut"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucDirectoryTree.ctx":0ECA
            Key             =   "FolderOpen"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucDirectoryTree.ctx":1276
            Key             =   "FolderOpenCut"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucDirectoryTree.ctx":161E
            Key             =   "MyComputer"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucDirectoryTree.ctx":19E2
            Key             =   "Desktop"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucDirectoryTree.ctx":1DBA
            Key             =   "Favourites"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucDirectoryTree.ctx":2192
            Key             =   "Unknown"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2990
      _Version        =   393217
      Indentation     =   353
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "DirectoryImages"
      Appearance      =   0
   End
   Begin VB.Menu mnuDirectory 
      Caption         =   "&Directory"
      Begin VB.Menu mnuCreate 
         Caption         =   "&New Folder..."
      End
      Begin VB.Menu mnuSeperatorA 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "&Rename..."
      End
      Begin VB.Menu mnuSeperatorC 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
   End
End
Attribute VB_Name = "ucDirectoryTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' ----------------------------------------------------------------------------------- '
' File........: ucDirectoryTree.ctl
' Version.....: 1.0.6
' Author......: Will Barden
' Created.....: 22/12/02
' Modified....: 10/01/03
'
' DirectoryTree control - for displaying the contents of a computer in
' treeview style. Uses a MS TreeView control to do the actual displaying,
' since writing a tree display would take ages.. :)
' Tight integration with Explorer and the Shell - uses SHFileOperation for
' all the file ops, and notifications from the shell when a directory is
' modified. In version 2, hopefully Cut/Copy/Paste integration as well.
' ----------------------------------------------------------------------------------- '
'
' ----------------------------------------------------------------------------------- '
' Events
' ----------------------------------------------------------------------------------- '
Public Event OnDirClick(ByVal sDir As String)
Public Event OnDirDblClick(ByVal sDir As String)
Public Event OnDirCreate(ByVal sDir As String)
Public Event OnDirRename(ByVal sOldPath As String, ByVal sNewPath As String)
Public Event OnDirRemove(ByVal sPath As String)
'
' ----------------------------------------------------------------------------------- '
' Private variables
' ----------------------------------------------------------------------------------- '
Private mbInitialized      As Boolean        ' whether we've been initialized or not.
Private mbSubclassed       As Boolean        ' whether we're subclassed or not.
Private mbRegistered       As Boolean        ' whether we're registered with the shell or not.
Private mlBorderType       As BorderTypes    ' Border style
Private mbHardDelete       As Boolean        ' should we delete or recycle files?
Private moCurNode          As Node           ' The currently selected directory node
'
' ----------------------------------------------------------------------------------- '
' Extra properties
' ----------------------------------------------------------------------------------- '
Public Property Get HardDelete() As Boolean
   '
   HardDelete = mbHardDelete
   '
End Property
'
Public Property Let HardDelete(ByVal Value As Boolean)
   '
   ' Turn on or off Hard Delete. If Value is True, then files are permanently
   ' removed, otherwise they're sent to the Recycle Bin.
   mbHardDelete = Value
   '
   PropertyChanged "HardDelete"
   '
End Property
'
Public Property Get BorderType() As BorderTypes
   '
   BorderType = mlBorderType
   '
End Property
'
Public Property Let BorderType(ByVal Value As BorderTypes)
   '
   ' Defines the style or border around the Usercontrol.
   mlBorderType = Value
   '
   PropertyChanged "BorderType"
   Call DoPaint
   '
End Property
'
Public Property Get Path() As String
   '
   If Not (moCurNode Is Nothing) Then
      Path = Mid$(moCurNode.Key, 2)
   End If
   '
End Property
'
Public Property Let Path(ByVal Value As String)
   '
   ' Create and show the specified directory :)
   Call CreateDirNode(Value, True, False)
   '
   RaiseEvent OnDirClick(Mid$(moCurNode.Key, 2))
   '
End Property
'
' ----------------------------------------------------------------------------------- '
' UserControl events
' ----------------------------------------------------------------------------------- '
Private Sub UserControl_Initialize()
   '
   ' Draw the control's border.
   Call DoPaint
   '
End Sub
'
Private Sub UserControl_Resize()
   '
   ' Fit the treeview to the UserControl. The magic numbers are there
   ' to make space for the control border (1 or 2 pixels wide).
   On Error Resume Next
   '
   With TreeView
      .Top = 2
      .Left = 2
      .Width = UserControl.ScaleWidth - 4
      .Height = UserControl.ScaleHeight - 4
   End With
   '
   ' Redraw the control.
   Call DoPaint
   '
End Sub
'
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   '
   ' Read the properties from the serialization.
   With PropBag
      mlBorderType = .ReadProperty("BorderType", BorderTypes.BDR_NONE)
      mbHardDelete = .ReadProperty("HardDelete", False)
   End With
   '
End Sub
'
Private Sub UserControl_Terminate()
   '
   ' The control is being shutdown - we have to detach ourselves from the
   ' various things.. firstly, the subclass.
   Call UnSubclass(UserControl.hWnd)
   '
   ' Then the shell change events.
   Call UnregisterNotify(UserControl.hWnd)
   '
End Sub
'
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   '
   ' Write the properties for serialization.
   With PropBag
      .WriteProperty "BorderType", mlBorderType, BorderTypes.BDR_NONE
      .WriteProperty "HardDelete", mbHardDelete, False
   End With
   '
End Sub
'
Private Sub UserControl_Show()
   '
   ' Call the custom paint.
   Call DoPaint
   '
End Sub
'
Private Sub UserControl_Paint()
   '
   ' Call the custom paint.
   Call DoPaint
   '
End Sub
'
' ----------------------------------------------------------------------------------- '
' Menu events
' ----------------------------------------------------------------------------------- '
Private Sub mnuCreate_Click()
Dim oNode   As Node
Dim i       As Long
Dim sTmp    As String
Dim sText   As String
   '
   ' Create a new folder under the currently selected node, and autostart
   ' the renaming process.
   '
   If (moCurNode Is Nothing) Then Exit Sub
   '
   ' Firstly, create the actual folder. To do this, we need to find a name
   ' that isn't in use already.
   sTmp = Mid$(moCurNode.Key, 2)
   sText = "New Folder"
   Do While LenB(Dir$(sTmp & sText, vbDirectory)) > 0
      sText = "New Folder (" & i & ")"
      i = i + 1
   Loop
   Call MkDir(sTmp & sText)
   '
   ' Add the new node to the tree, with an appropriate key.
   sTmp = UCase$(AddSlash(FLDR_NORMAL & sTmp & sText))
   Set oNode = TreeView.Nodes.Add(moCurNode, _
                                  tvwChild, _
                                  sTmp, _
                                  sText, _
                                  "Folder", _
                                  "FolderOpen")
   '
   ' Select, and then start the renaming process.
   Set moCurNode = oNode
   moCurNode.Selected = True
   Call TreeView.StartLabelEdit
   '
   RaiseEvent OnDirCreate(Mid$(moCurNode.Key, 2))
   '
End Sub
'
Private Sub mnuDelete_Click()
Dim eRet    As VbMsgBoxResult
Dim lRet    As Long
Dim lFlags  As Long
Dim uFileOp As SHFILEOP_STRUCT
Dim sPath   As String
   '
   ' Make sure we've got a selected node.
   If (moCurNode Is Nothing) Then Exit Sub
   sPath = Mid$(moCurNode.Key, 2, Len(moCurNode.Key) - 2)
   '
   ' Check the HardDelete property - if it's true, then files are permanently
   ' removed, otherwise they're moved to the Recycle Bin. Both operations are
   ' done via the SHFileOperation API call - this provides the nice Windows
   ' progress dialog for long operations :).
   '
   If mbHardDelete Then
      ' Set the flags to 0 - permanently delete.
      lFlags = 0
   Else
      ' Set the flags to allow undo - move to the recycle bin.
      lFlags = FOF_ALLOWUNDO
   End If
   '
   ' Setup the file operation information struct, and delete.
   With uFileOp
      .wFunc = FO_DELETE
      .fFlags = lFlags
      .pFrom = sPath
   End With
   lRet = ApiSHFileOperation(uFileOp)
   '
   If lRet = 0 Then
      '
      ' The deletion was a success, but we don't need to remove the node, since it
      ' will already have been done when we received a notification from the shell.
      Set moCurNode = Nothing
      '
      RaiseEvent OnDirRemove(sPath)
      '
   End If
   '
End Sub
'
Private Sub mnuRename_Click()
   '
   ' Show the label edit for the directory, and then record the results,
   ' renaming the directory as appropriate.
   If Not (moCurNode Is Nothing) Then
      Call TreeView.StartLabelEdit
   End If
   '
End Sub
'
' ----------------------------------------------------------------------------------- '
' Treeview events
' ----------------------------------------------------------------------------------- '
Private Sub TreeView_BeforeLabelEdit(Cancel As Integer)
Dim sTmp As String
   '
   ' Check that node is a real folder, then allow the rename.
   If (moCurNode Is Nothing) Then Exit Sub
   '
   sTmp = Left$(moCurNode.Key, 1)
   '
   If (sTmp = FLDR_VIRTUAL) Or (sTmp = FLDR_DRIVE) Then
      Cancel = True
   End If
   '
End Sub
'
Private Sub TreeView_AfterLabelEdit(Cancel As Integer, NewString As String)
Dim i          As Long
Dim sNewName   As String
   '
   ' If it's not been changed, don't bother continuing.
   If (moCurNode.Text = NewString) Or (LenB(NewString) = 0) Then
      Cancel = True
      Exit Sub
   End If
   '
   ' Validate the new directory name for invalid chars, and then rename.
   ' \/:*?"<>| - invalid folder name characters
   For i = 1 To Len(FLDR_INVALID_CHARS)
      If InStr(1, NewString, Mid$(FLDR_INVALID_CHARS, i, 1)) Then
         '
         ' NewString is an invalid folder name
         MsgBox "Invalid directory name. Directory names may not contain the" & vbCrLf & _
                "following characters: " & FLDR_INVALID_CHARS & Chr$(34), vbOKOnly + vbCritical
         Cancel = True
         Exit Sub
         '
      End If
   Next i
   '
   ' Check for Chr$(34) (a quote char)
   If InStr(1, NewString, Chr$(34)) Then
      '
      ' NewString is an invalid folder name
      MsgBox "Invalid directory name. Directory names may not contain the" & vbCrLf & _
             "following characters: " & FLDR_INVALID_CHARS & Chr$(34), vbOKOnly + vbCritical
      Cancel = True
      Exit Sub
      '
   End If
   '
   ' If we've got here, then all the characters are ok, so rename the folder.
   ' Replace the last folder in the string with the NewString (new folder name).
   sNewName = Mid$(moCurNode.Key, 2, InStrRev(moCurNode.Key, "\", Len(moCurNode.Key) - 1) - 1)
   sNewName = sNewName & NewString & "\"
   Name Mid$(moCurNode.Key, 2) As sNewName
   '
   ' Now alter the node's key to reflect the change.
   moCurNode.Key = UCase$(AddSlash(FLDR_NORMAL & sNewName))
   '
   RaiseEvent OnDirRename(Mid$(moCurNode.Key, 2), sNewName)
   '
End Sub
'
Private Sub TreeView_Expand(ByVal Node As MSComctlLib.Node)
Dim i       As Long
Dim oNode   As Node
   '
   On Error GoTo Errored
   '
   ' A node that we know has children is being expanded - for each child
   ' directory, find all the child's child directories, and add them.
   Set oNode = Node.Child
   For i = 1 To Node.Children
      '
      ' If this child has no children, then search it, since we've not
      ' done it already.
      If oNode.Children = 0 Then
         Call ShowChildren(TreeView.Nodes(oNode.Index), False)
      End If
      '
      ' Get the next child node in the list.
      Set oNode = oNode.Next
      '
   Next i
   Exit Sub
   '
Errored:
   '
   ' Something went wrong somewhere..
   If Err.Number = DirectoryErrors.DERR_NO_DIRECTORY_LIST Then
      MsgBox "Folder or drive not ready: " & Err.Description, vbOKOnly + vbCritical
   Else
      MsgBox Err.Number & ": " & Err.Description, vbOKOnly + vbCritical
   End If
   '
   Resume Next
   '
End Sub
'
Private Sub TreeView_NodeClick(ByVal Node As MSComctlLib.Node)
   '
   ' We should store a reference to the currently selected node.
   Set moCurNode = Node
   '
   ' Since a directory has been clicked, raise the event to the parent.
   RaiseEvent OnDirClick(Mid$(moCurNode.Key, 2))
   '
End Sub
'
Private Sub TreeView_DblClick()
   '
   On Error GoTo Errored
   '
   ' Raise the OnDirDblClick event if a directory is selected.
   If Not (moCurNode Is Nothing) Then
      '
      ' If this node has no children (for example, a removable media drive),
      ' show them now.
      If (moCurNode.Children = 0) Then
         Call ShowChildren(moCurNode, True)
      End If
      '
      RaiseEvent OnDirDblClick(Mid$(moCurNode.Key, 2))
      '
   End If
   Exit Sub
   '
Errored:
   '
   ' Something went wrong somewhere..
   If Err.Number = DirectoryErrors.DERR_NO_DIRECTORY_LIST Then
      MsgBox "Folder or drive not ready: " & Err.Description, vbOKOnly + vbCritical
   Else
      MsgBox Err.Number & ": " & Err.Description, vbOKOnly + vbCritical
   End If
   '
   Resume Next
   '
End Sub
'
Private Sub TreeView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim sType As String
   '
   ' On a right click, we want to display our menu of options.
   If Button = vbRightButton Then
      '
      ' Check that we have a node currently selected.
      If Not (moCurNode Is Nothing) Then
         '
         ' Re-enable the other menu items, in case they've been disabled,
         ' BUT - only enable if the current node is not a virtual folder;
         ' and, since we can't delete drives, disable them too.
         sType = Left$(moCurNode.Key, 1)
         If sType = FLDR_VIRTUAL Or sType = FLDR_DRIVE Then
            mnuCreate.Enabled = False
            mnuRename.Enabled = False
            mnuDelete.Enabled = False
         Else
            mnuCreate.Enabled = True
            mnuRename.Enabled = True
            mnuDelete.Enabled = True
         End If
         '
      Else
         '
         ' Since we don't have a directory selected, disable the entire menu.
         mnuCreate.Enabled = False
         mnuRename.Enabled = False
         mnuDelete.Enabled = False
         '
      End If
      '
      ' Finally, popup the menu :)
      Call PopupMenu(mnuDirectory)
      '
   End If
   '
End Sub
'
' ----------------------------------------------------------------------------------- '
' Methods
' ----------------------------------------------------------------------------------- '
Public Sub Clear()
   '
   ' Clear the entire Treeview control - the user may wish to do this in order
   ' to refresh the list of drives shown.
   Call TreeView.Nodes.Clear
   mbInitialized = False
   '
End Sub
'
Public Sub Initialize()
   '
   ' There's quite a lot to do, so do it all at once. Initially, check that
   ' we haven't been setup already.
   If mbInitialized Then
      '
      Err.Raise DirectoryErrors.DERR_ALREADY_INITIALIZED, _
                "ucDirectoryTree.Initialize", _
                "Tree already initialized."
      Exit Sub
      '
   End If
   '
   ' Firstly, subclass the control, so that we can start receiving shell
   ' change events. This is very important.
   If Not mbSubclassed Then
      '
      mbSubclassed = Subclass(UserControl.hWnd, ObjPtr(Me))
      If Not mbSubclassed Then
         '
         ' Failed to subclass the control, so let the user know. Should we
         ' terminate at this point?
         Err.Raise DirectoryErrors.DERR_SUBCLASS_FAILED, _
                   "ucDirectoryTree.Initialize", _
                   "Couldn't subclass tree."
         Exit Sub
         '
      End If
      '
   End If
   '
   ' Next, we need to register the control as a receiver for shell change
   ' notification, so we can update our tree when someone renames, creates or
   ' removes a folder. (moves?)
   If Not mbRegistered Then
      '
      mbRegistered = RegisterNotify(UserControl.hWnd)
      If Not mbRegistered Then
         '
         ' Failed to register for shell change notification event. Pants.
         Err.Raise DirectoryErrors.DERR_REGISTER_FAILED, _
                   "ucDirectoryTree.Initialize", _
                   "Couldn't register for shell change events"
         Exit Sub
         '
      End If
      '
   End If
   '
   ' Lastly, we need to fill up the tree with our root nodes. Nice and easy.
   Call Clear
   Call ShowRoots
   '
   mbInitialized = True
   '
End Sub
'
' ----------------------------------------------------------------------------------- '
' Callback routine from the mGeneral.bas module.
' ----------------------------------------------------------------------------------- '
Public Sub ShellNotifyMsg(ByVal wParam As Long, _
                          ByVal lParam As Long)
Dim uNotify As SHNOTIFY_EVENT
Dim sOldPath  As String
Dim sNewPath  As String
   '
   ' We have received a shell change notification!! Woooo!!
   ' wParam is a pointer to a SHNOTIFY_EVENT, and lParam is the message
   ' that we've received.
   '
   ' So, first - get the SHNOTIFY_EVENT and extract the two paths (old and
   ' new in the case of a renaming, or deleting/moving; for creations, new will
   ' be blank).
   Call ApiCopyMemory(uNotify, ByVal wParam, LenB(uNotify))
   sOldPath = GetPathFromPidl(uNotify.dwItem1)
   sNewPath = GetPathFromPidl(uNotify.dwItem2)
   '
   ' Now check the message we've been sent to react to the shell event.
   Select Case lParam
      '
      ' A directory has been created.
      Case SHCNE_MKDIR
         Call DirectoryCreated(sOldPath)
         RaiseEvent OnDirCreate(sOldPath)
      '
      ' A directory has been removed.
      Case SHCNE_RMDIR
         Call DirectoryRemoved(sOldPath, sNewPath)
         RaiseEvent OnDirRemove(sOldPath)
      '
      ' A directory has been renamed/moved.
      Case SHCNE_RENAMEFOLDER
         Call DirectoryRenamed(sOldPath, sNewPath)
         RaiseEvent OnDirRename(sOldPath, sNewPath)
      '
   End Select
   '
End Sub
'
' ----------------------------------------------------------------------------------- '
' Private helper routines
' ----------------------------------------------------------------------------------- '
Private Sub ShowRoots()
Dim i          As Long
Dim lRet       As Long
Dim sBuffer    As String
Dim sDrives()  As String
Dim lDriveType As DriveTypes
Dim sIcon      As String
Dim oNode      As Node
Dim sDesktop   As String
Dim sFav       As String
Dim pIdl       As Long
   '
   ' Retrieve and display a list of available drives on the system, including
   ' network, CDROM, RAMDISK, floppy, and HDD.
   '
   ' Add the root node - MyComputer.
   Set oNode = TreeView.Nodes.Add(, , FLDR_VIRTUAL & "MyComputer", "My Computer", "MyComputer")
   oNode.Expanded = True
   '
   ' Prepare a string buffer to retrieve the drive list. This has to be sufficiently
   ' large in order to receive all available drives - if we imagine that four characters
   ' are required for each drive, BUFFER_LEN is enough for approx 250 drives, plenty.
   sBuffer = String$(BUFFER_LEN, 0)
   lRet = ApiGetLogicalDriveStrings(Len(sBuffer), sBuffer)
   '
   ' Check the return value - this is the number of bytes copied into the buffer.
   If lRet = 0 Then
      '
      ' Failed to retrieve the list, so raise an error.
      Err.Raise DirectoryErrors.DERR_NO_DRIVE_LIST, _
                "ucDirectoryTree.ShowRoots", _
                "Couldn't retrieve drive list"
      Exit Sub
      '
   Else
      '
      ' Trim the drive list to the proper length, and split into an array.
      sBuffer = Left$(sBuffer, lRet - 1)
      sDrives = Split(sBuffer, Chr$(0))
      '
      ' For each drive, get it's type, and prepare the image to be used.
      For i = LBound(sDrives) To UBound(sDrives)
         '
         lDriveType = ApiGetDriveType(sDrives(i))
         Select Case lDriveType
            Case DRIVE_UNKNOWN, DRIVE_NO_ROOT_DIR
               sIcon = "Unknown"
            Case DRIVE_REMOVABLE
               sIcon = "Removable"
            Case DRIVE_FIXED
               sIcon = "Fixed"
            Case DRIVE_REMOTE
               sIcon = "Remote"
            Case DRIVE_CDROM
               sIcon = "CdRom"
            Case DRIVE_RAMDISK
               sIcon = "RamDisk"
            Case Else            ' default
               sIcon = "Unknown"
         End Select
         '
         ' Now add the item as a root node to the treeview, setting the drive letter
         ' as the key - also, set the appropriate icon from the drive type.
         Set oNode = TreeView.Nodes.Add(FLDR_VIRTUAL & "MyComputer", tvwChild, FLDR_DRIVE & UCase$(AddSlash(sDrives(i))), UCase$(sDrives(i)), sIcon)
         '
         ' Show this drives top level children - only do this if it's a fixed
         ' disk to save time - sometimes removable or CD/DVD drives can be a bit
         ' slow, and we don't want to incur unwanted slowness.
         If lDriveType <> DRIVE_REMOVABLE And lDriveType <> DRIVE_CDROM Then
            Call ShowChildren(oNode, False)
         End If
         '
      Next i
      '
   End If
   '
End Sub
'
Private Function NodeExists(ByVal sKey As String) As Long
Dim oTmp    As Node
   '
   ' Check to see if a node exists in our collection by trying to
   ' access it. If it doesn't exist, an error will fire. Returns the
   ' node's Index.
   On Error Resume Next
   Call Err.Clear
   '
   ' Now check for it.
   Set oTmp = TreeView.Nodes(sKey)
   If Err.Number = 0 Then
      NodeExists = oTmp.Index
   End If
   '
End Function
'
Private Sub ClearChildren(ByRef oNode As Node)
Dim i As Long
   '
   ' Recursively travel the branch under this node, and clear all the
   ' child nodes.
   For i = 1 To oNode.Children
      '
      If oNode.Child.Children > 0 Then
         Call ClearChildren(oNode)
      End If
      '
      Call TreeView.Nodes.Remove(oNode.Child.Index)
      '
   Next i
   '
End Sub
'
Private Sub ShowChildren(ByRef oNode As Node, ByVal bExpand As Boolean)
Dim sDir    As String
Dim hFile   As Long
Dim uData   As WIN32_FINDDATA
Dim sTmp    As String
Dim lRet    As Long
Dim sKey    As String
   '
   ' Trim out the folder description character at the beginning - this is
   ' not used for anything except avoiding key collisions so far, but may
   ' become useful in later versions.
   sDir = Mid$(oNode.Key, 2)
   '
   ' either FLDR_VIRTUAL: MyComputer, or FLDR_SPECIAL: desktop,
   ' We have been asked to show the child directories of a parent node.
   ' Use the file APIs to travel the top level of that directory, and
   ' add any directories as children under this node.
   hFile = ApiFindFirstFile(sDir & FIND_ALL_FILES, uData)
   If hFile = INVALID_HANDLE_VALUE Then
      '
      ' Failed to initialize the search..
      Err.Raise DirectoryErrors.DERR_NO_DIRECTORY_LIST, _
                "ucDirectoryTree.ShowChildren", _
                "Failed to open search handle to: " & sDir
      Exit Sub
      '
   End If
   '
   ' Close this node up, and remove the sorted property - this will
   ' make adding nodes a lot faster.
   With oNode
      .Expanded = False
      .Sorted = False
   End With
   '
   ' Start the loop, to continue through the entire directory
   Do
      '
      ' The handle is valid, so the WIN32_FINDDATA struct contains into
      ' about the first item located. Since we only want directories, check to
      ' make sure that it is one.
      If uData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
         '
         ' Trim off the filename, and check it's not a navigational control
         ' - "." or ".."
         sTmp = uData.cFileName
         Call TrimNulls(sTmp)
         If sTmp <> "." And sTmp <> ".." Then
            '
            ' We have a valid directory name, so add it as a child node.
            sKey = Mid$(oNode.Key, 2)
            TreeView.Nodes.Add oNode, _
                               tvwChild, _
                               UCase$(AddSlash(FLDR_NORMAL & sKey & sTmp)), _
                               sTmp, _
                               "Folder", _
                               "FolderOpen"
            '
         End If
         '
      End If
      '
      ' Find the next item in the directory, whether it be a file or sub-dir.
      lRet = ApiFindNextFile(hFile, uData)
      '
   Loop While lRet <> 0
   '
   ' We must always closeup the file search handle once we're done.
   Call ApiFindClose(hFile)
   '
   ' Finally, expand the node we've just filled, so that the children are
   ' on display, and sort them alphabetically.
   With oNode
      .Sorted = True
      .Expanded = bExpand
   End With
   '
End Sub
'
Private Sub DirectoryCreated(ByVal sDir As String)
   '
   ' A new directory has been created - we know the path, so add it in to
   ' the tree - BUT, only do this if we've got it's parent added in.
   Call CreateDirNode(sDir, False, True)
   '
End Sub
'
Private Sub DirectoryRemoved(ByVal sOldDir As String, _
                             ByVal sNewDir As String)
Dim lNodeIndex As Long
   '
   ' A directory has been removed from disk, into the Recycle Bin (sNewDir
   ' contains the exact path within the Bin). All we have to do is remove
   ' the node, if we've got it.
   lNodeIndex = NodeExists(UCase$(AddSlash(FLDR_NORMAL & sOldDir)))
   If lNodeIndex Then
      '
      Call TreeView.Nodes.Remove(lNodeIndex)
      '
   End If
   '
End Sub
'
Private Sub DirectoryRenamed(ByVal sOldDir As String, _
                             ByVal sNewDir As String)
Dim lNodeIndex    As Long
Dim sChildDir     As String
Dim sOldParentDir As String
Dim sNewParentDir As String
Dim oNode         As Node
   '
   ' A directory somewhere has been renamed or moved. If we have the old
   ' directory in our tree, then we should alter it.
   lNodeIndex = NodeExists(UCase$(AddSlash(FLDR_NORMAL & sOldDir)))
   If lNodeIndex Then
      '
      ' We must now check whether this is a rename, or a move, because the
      ' operation will be different for each. To do this, strip off the last
      ' directory, and then compare the parent paths. If they are different,
      ' then the child dir has been moved.
      '
      sOldParentDir = Mid$(sOldDir, 1, InStrRev(sOldDir, "\", Len(sOldDir) - 1))
      sNewParentDir = Mid$(sNewDir, 1, InStrRev(sNewDir, "\", Len(sNewDir) - 1))
      sOldParentDir = UCase$(AddSlash(sOldParentDir))
      sNewParentDir = UCase$(AddSlash(sNewParentDir))
      '
      If sOldParentDir <> sNewParentDir Then
         '
         ' A move!! TO deal with this we'll have to do a couple of things.
         ' Firstly, and most importantly, remove the old directory node, since
         ' it no longer exists.
         Call TreeView.Nodes.Remove(lNodeIndex)
         '
         ' Second, add the new directory node. We only want to bother doing this
         ' if the new directory's parent already exists.
         If NodeExists(FLDR_NORMAL & sNewParentDir) Then
            Call CreateDirNode(sNewDir, False, True)
         End If
         '
      Else
         '
         ' Just a rename. Easy - we have the node's index, just alter the key
         ' and the text accordingly.
         Set oNode = TreeView.Nodes(lNodeIndex)
         With oNode
            '
            sChildDir = Mid$(sNewDir, InStrRev(sNewDir, "\", Len(sNewDir) - 1) + 1)
            sChildDir = Left$(sChildDir, Len(sChildDir) - 1)
            '
            .Text = sChildDir
            .Key = UCase$(AddSlash(Mid$(.Key, 1, 1) & sNewParentDir & sChildDir))
            '
         End With
         '
         ' Re-sort the parent to re-place the nodes in the right order. :)
         oNode.Parent.Sorted = True
         '
      End If
      '
   End If
   '
End Sub
'
Private Sub CreateDirNode(ByVal sDir As String, _
                          ByVal bExpand As Boolean, _
                          ByVal bCreate As Boolean)
Dim lPos    As Long
Dim sPath   As String
Dim sTmp    As String
Dim lIndex  As Long
Dim oNode   As Node
Dim sText   As String
   '
   ' Start by normalizing the path, just in case the user has given us
   ' a dodgy one.
   sPath = UCase$(AddSlash(sDir))
   '
   ' Now strip out the drive name, and make sure we have it loaded.
   lPos = InStr(1, sPath, "\")
   If lPos = 0 Then
      '
      ' No slashes?!?
      Exit Sub
      '
   End If
   sTmp = UCase$(Mid$(sPath, 1, lPos))
   lIndex = NodeExists(FLDR_DRIVE & sTmp)
   If lIndex Then
      '
      ' Now we know the drive exists, get all of it's children.
      Set moCurNode = TreeView.Nodes(lIndex)
      If moCurNode.Children = 0 Then Call ShowChildren(moCurNode, bExpand)
      '
      ' Go through the specified path, and select each of the children in turn.
      Do
         '
         lPos = InStr(lPos + 1, sPath, "\")
         If lPos > 0 Then
            '
            ' Get the folder name and check it exists on the local drive.
            sTmp = Mid$(sPath, 1, lPos)
            If LenB(Dir$(sTmp, vbDirectory)) = 0 Then
               '
               ' Folder doesn't exist!
               Err.Raise DirectoryErrors.DERR_INVALID_FOLDER, _
                         "ucDirectoryTree.CreateDirNode", _
                         "Specified folder: """ & sTmp & """ doesn't exist."
               Exit Sub
               '
            End If
            '
            ' Check we have it in the tree.
            lIndex = NodeExists(FLDR_NORMAL & sTmp)
            If lIndex Then
               '
               ' We have a directory that exists, and that exists in our tree.
               ' Load all it's children, ready for the next sub-dir.
               Set moCurNode = TreeView.Nodes(lIndex)
               If moCurNode.Children = 0 Then Call ShowChildren(moCurNode, bExpand)
               '
            Else
               '
               ' The specified directory exists on disk, but not in our tree. As such,
               ' create the directory if we've been told to, then show it's children.
               If bCreate Then
                  '
                  ' Get a reference to the parent node, and split off the new
                  ' child's name for the text property.
                  lIndex = NodeExists(FLDR_NORMAL & Mid$(sTmp, 1, InStrRev(sTmp, "\", Len(sTmp) - 1)))
                  If lIndex Then
                     Set oNode = TreeView.Nodes(lIndex)
                     '
                     sText = Mid$(sDir, InStrRev(sDir, "\", Len(sDir) - 1) + 1)
                     sText = Mid$(sText, 1, Len(sText) - 1)
                     '
                     Set oNode = TreeView.Nodes.Add(oNode, _
                                                    tvwChild, _
                                                    FLDR_NORMAL & sTmp, _
                                                    sText, _
                                                    "Folder", _
                                                    "FolderOpen")
                     Call ShowChildren(oNode, bExpand)
                     '
                  End If
                  '
               Else
                  '
                  ' This directory doesn't exist in the tree...
                  Err.Raise DirectoryErrors.DERR_INVALID_FOLDER, _
                            "ucDirectoryTree.CreateDirNode", _
                            "Directory not loaded in tree: """ & sTmp & """."
                  Exit Sub
                  '
               End If
               '
            End If
            '
         End If
         '
      Loop While lPos > 0 And lPos <= Len(sPath)
      '
   Else
      '
      ' The drive doesn't exist, so raise an error.
      Err.Raise DirectoryErrors.DERR_INVALID_DRIVE, _
                "ucDirectoryTree.CreateDirNode", _
                "Specified drive doesn't exist in the tree."
      Exit Sub
      '
   End If
   '
   ' Finally, select and then raise a click event on the selected dir.
   With moCurNode
      .Selected = bExpand
      If bExpand Then
         .Expanded = bExpand
      End If
   End With
   Exit Sub
   '
End Sub
'
Private Sub DoPaint()
Dim uRect As RECT
   '
   ' Setup the rectangle to draw the border onto.
   With uRect
      .Right = UserControl.ScaleWidth
      .Bottom = UserControl.ScaleHeight
   End With
   '
   ' Draw a border onto the UserControl, just around the Treeview.
   With UserControl
      '
      Call .Cls
      Call ApiDrawEdge(.hdc, uRect, mlBorderType, BF_RECT)
      '
   End With
   '
End Sub
'
