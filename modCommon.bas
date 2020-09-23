Attribute VB_Name = "modCommon"
Option Explicit

Public Declare Function DoAddToFavDlg Lib "shdocvw.dll" _
  (ByVal hwnd As Long, _
   ByVal szPath As String, _
   ByVal nSizeOfPath As Long, _
   ByVal szTitle As String, _
   ByVal nSizeOfTitle As Long, _
   ByVal pidl As Long) As Long
   
Public Declare Function DoOrganizeFavDlg Lib "shdocvw.dll" _
  (ByVal hwnd As Long, _
   ByVal lpszRootFolder As String) As Long


Public Const MAXDWORD = &HFFFFFFF
Public Const MAX_PATH = 260
Public Const INVALID_HANDLE_VALUE = -1
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const FILE_ATTRIBUTE_FLAGS = FILE_ATTRIBUTE_ARCHIVE Or _
                                     FILE_ATTRIBUTE_HIDDEN Or _
                                     FILE_ATTRIBUTE_NORMAL Or _
                                     FILE_ATTRIBUTE_READONLY

Public Const SHGFP_TYPE_CURRENT As Long = &H0
Public Const SHGFP_TYPE_DEFAULT As Long = &H1
Public Const CSIDL_FAVORITES As Long = &H6
Public Const CSIDL_COMMON_FAVORITES As Long = &H1F
Public Const MAX_LENGTH As Long = 260
Public Const S_OK As Long = 0
Public Const S_FALSE As Long = 1


Public Const DRIVE_UNKNOWNTYPE = 1
Public Const DRIVE_REMOVABLE = 2
Public Const DRIVE_FIXED = 3
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_CDROM = 5
Public Const DRIVE_RAMDISK = 6

Public Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

'the custom-made UDT for searching
Public Type FILE_PARAMS
   bRecurse As Boolean     'set True to perform a recursive search
   bList As Boolean        'set True to add results to listbox
   bFound As Boolean       'set only with SearchTreeForFile methods
   sFileRoot As String     'search starting point, ie c:\, c:\winnt\
   sFileNameExt As String  'filename/filespec to locate, ie *.dll, notepad.exe
   sResult As String       'path to file. Set only with SearchTreeForFile methods
   nFileCount As Long      'total file count matching filespec. Set in FindXXX only
   nFileSize As Double     'total file size matching filespec. Set in FindXXX only
End Type

Public Declare Function FindClose Lib "kernel32" _
  (ByVal hFindFile As Long) As Long
   
Public Declare Function FindFirstFile Lib "kernel32" _
   Alias "FindFirstFileA" _
  (ByVal lpFileName As String, _
   lpFindFileData As WIN32_FIND_DATA) As Long
   
Public Declare Function FindNextFile Lib "kernel32" _
   Alias "FindNextFileA" _
  (ByVal hFindFile As Long, _
   lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function GetLogicalDriveStrings Lib "kernel32" _
   Alias "GetLogicalDriveStringsA" _
  (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
      
Public Declare Function GetDriveType Lib "kernel32" _
   Alias "GetDriveTypeA" _
  (ByVal nDrive As String) As Long

Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public Const SW_SHOWNA = 8
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWDEFAULT = 10


Public Declare Function ShellExecute Lib "shell32.dll" _
   Alias "ShellExecuteA" _
  (ByVal hwnd As Long, ByVal lpOperation As String, _
   ByVal lpFile As String, ByVal lpParameters As String, _
   ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function FindExecutable Lib "shell32.dll" Alias _
         "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As _
         String, ByVal lpResult As String) As Long
   

Public Declare Function GetPrivateProfileString _
   Lib "kernel32" Alias "GetPrivateProfileStringA" _
  (ByVal lpSectionName As String, _
   ByVal lpKeyName As Any, _
   ByVal lpDefault As String, _
   ByVal lpReturnedString As String, _
   ByVal nSize As Long, _
   ByVal lpFileName As String) As Long
   
Public Declare Function SHGetFolderPath _
    Lib "shfolder.dll" Alias "SHGetFolderPathA" _
   (ByVal hWndOwner As Long, _
     ByVal nFolder As Long, _
     ByVal hToken As Long, _
     ByVal dwReserved As Long, _
     ByVal lpszPath As String) As Long
      

Type FavMenu
    FavName As String
    FavURL As String
    FavRoot As String
    FavFolder As String
End Type


Private Sub GetFileInformation(FP As FILE_PARAMS)
On Error Resume Next
  'local working variables
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   Dim sPath As String
   Dim sRoot As String
   Dim sTmp As String

  'FP.sFileRoot (assigned to sRoot) contains
  'the path to search.
  
  'FP.sFileNameExt (assigned to sPath) contains
  'the full path and filespec.
   sRoot = QualifyPath(FP.sFileRoot)
   sPath = sRoot & FP.sFileNameExt

  'obtain handle to the first filespec match
   hFile = FindFirstFile(sPath, WFD)

  'if valid ...
   If hFile <> INVALID_HANDLE_VALUE Then

      Do

        'Even though this routine uses filespecs,
        '*.* is still valid and will cause the search
        'to return folders as well as files, so a
        'check against folders is still required.
         If Not (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) _
                 = FILE_ATTRIBUTE_DIRECTORY Then

           'remove trailing nulls
            sTmp = TrimNull(WFD.cFileName)

         End If

      Loop While FindNextFile(hFile, WFD)

     'close the handle
      hFile = FindClose(hFile)

   End If

End Sub


Private Sub GetAllFilesSpecified(FP As FILE_PARAMS)
On Error Resume Next
   Dim drvCount As Long
   Dim sBuffer As String
   Dim currDrive As String
   
   If FP.sFileRoot = "all fixed disks/partitions" Then
   
     'all drives
   
     'retrieve the available drives on the system
      sBuffer = Space$(64)
      drvCount = GetLogicalDriveStrings(Len(sBuffer), sBuffer)
   
     'drvCount returns the size of the drive string
      If drvCount Then
      
        'strip off trailing nulls
         sBuffer = Left$(sBuffer, drvCount)
              
        'search each drive for the file
         Do Until sBuffer = ""
   
           'strip off one drive item from sBuffer
            FP.sFileRoot = StripItem(sBuffer)
   
           'just search the local file system
            If GetDriveType(FP.sFileRoot) = DRIVE_FIXED Then
            
                             
               DoEvents
                              
               Call SearchForFilesArray(FP)
               
            End If
         
         Loop
      
      End If
      
   Else
   
       
       Call SearchForFilesArray(FP)
       
   End If

End Sub


Public Function TrimNull(startstr As String) As String
On Error Resume Next
  'returns the string up to the first
  'null, if present, or the passed string
   Dim pos As Integer
   
   pos = InStr(startstr, Chr$(0))
   
   If pos Then
      TrimNull = Left$(startstr, pos - 1)
      Exit Function
   End If
  
   TrimNull = startstr
  
End Function


Private Sub SearchForFilesArray(FP As FILE_PARAMS)
On Error Resume Next
  'local working variables
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   Dim sPath As String
   Dim sRoot As String
   Dim sTmp As String
      
  'this routine is primarily interested in the
  'directories, so the file type must be *.*
   sRoot = QualifyPath(FP.sFileRoot)
   sPath = sRoot & "*.*"
   
  'obtain handle to the first match
   hFile = FindFirstFile(sPath, WFD)
   
  'if valid ...
   If hFile <> INVALID_HANDLE_VALUE Then
   
     'GetFileInformation function returns the number,
     'of files matching the filespec (FP.sFileNameExt)
     'in the passed folder.
      Call GetFileInformation(FP)

      Do
      
        'if the returned item is a folder...
         If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
            
           'remove trailing nulls
            sTmp = TrimNull(WFD.cFileName)
            
           'and if the folder is not the default
           'self and parent folders...
            If sTmp <> "." And sTmp <> ".." Then
            
              'get the file
               FP.sFileRoot = sRoot & sTmp
               
               
              'This next If..Then just prevents adding extra
              'lines and unneeded paths to the array when a
              'file search is performed for a specific file type
               If FP.sFileNameExt = "*.*" Then
               
                 'Depending on the purpose, you may want to
                 'exclude the next 4 optional lines.
                 'The first two lines adds a blank entry
                 'to the array as a separator between new
                 'directories in the output file. The last
                 'two add the directory name alone, before
                 'listing the files underneath. These four
                 'lines can be optionally commented out).
                 'Obviously, these extra entries will skew
                 'the actual file counts.
                                 
               End If
               
              'call again
               Call SearchForFilesArray(FP)
            
            End If
               
            
         End If
         
     'continue looping until FindNextFile returns
     '0 (no more matches)
      Loop While FindNextFile(hFile, WFD)
      
     'close the find handle
      hFile = FindClose(hFile)
   
   End If
   
End Sub


Private Function QualifyPath(sPath As String) As String
On Error Resume Next
  'assures that a passed path ends in a slash
   If Right$(sPath, 1) <> "\" Then
         QualifyPath = sPath & "\"
   Else: QualifyPath = sPath
   End If
      
End Function


Function StripItem(startStrg As String) As String
On Error Resume Next
  'Take a string separated by Chr(0)'s, and split off 1 item, and
  'shorten the string so that the next item is ready for removal.
   Dim pos As Integer

   pos = InStr(startStrg, Chr$(0))

   If pos Then
      StripItem = Mid(startStrg, 1, pos - 1)
      startStrg = Mid(startStrg, pos + 1, Len(startStrg))
   End If

End Function



Private Sub GetSystemDrives(ctl As ComboBox)
On Error Resume Next
   Dim drvCount As Long
   Dim sBuffer As String
   Dim currDrive As String
       
  'Retrieve the available drives on the system.
  'The string is padded with enough room to hold
  'the drives, nulls and final terminating null.
   sBuffer = Space$(105)
   drvCount = GetLogicalDriveStrings(Len(sBuffer), sBuffer)
   
  'drvCount returns the size of the drive string
   If drvCount Then
   
     'strip off trailing nulls
      sBuffer = Left$(sBuffer, drvCount)
           
     'search each drive for the file
      Do Until sBuffer = ""

        'strip off one drive item from sBuffer
         currDrive = StripItem(sBuffer)

        'just search the local file system
         If GetDriveType(currDrive) = DRIVE_FIXED Then
         
            ctl.AddItem Left$(currDrive, 2)
            
         End If
      
      Loop
      
   End If

End Sub

Public Sub RunShellExecute(sTopic As String, _
                            sFIle As Variant, _
                            sParams As Variant, _
                            sDirectory As Variant, _
                            nShowCmd As Long)
On Error Resume Next
   Dim hWndDesk As Long
   Dim success As Long
  
  'the desktop will be the
  'default for error messages
   hWndDesk = GetDesktopWindow()
  
  'execute the passed operation
   success = ShellExecute(hWndDesk, sTopic, sFIle, sParams, sDirectory, nShowCmd)

  'have the "Open With.." dialog appear
  'when the ShellExecute API call fails
  If success < 32 Then
     Call Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sFIle, vbNormalFocus)
  End If
   
End Sub

