Attribute VB_Name = "File_search"
'From VBhelp
Option Explicit

Private Const SW_SHOWMAXIMIZED = 3
Private Const ArrGrow As Long = 5000
Private Const MaxLong As Long = 2147483647
Private Const MAX_PATH = 260
Private Const MAXDWORD = &HFFFF
Private Const INVALID_HANDLE_VALUE = -1
Private Const LB_SETHORIZONTALEXTENT = &H194
Private Const LB_ADDSTRING = &H180



Enum eSortMethods
    SortNot = 0
    SortByNames = 1
End Enum

Enum eSizeConstants
    BIPerB = 8
    BPERKB = 1024
    KBPerMB = 1024
    MBPerGB = 1024
    GBPerTB = 1024
    TBPerPT = 1024
End Enum

Private Type TextSize
    Width As Long
    Height As Long
End Type

Type tFile
    Name As String
    Path As String
    FullName As String
    CreationDate As String
    AccessDate As String
    WriteDate As String
    Size As Currency
    Attr As VbFileAttribute
End Type

Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved As Long
    dwReserved1 As Long
    FileName As String * MAX_PATH
    cAlternateFileName As String * 14
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

'Window
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function GetFocus Lib "user32" () As Long

'Shell
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long
Private Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

'File Stuff
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'Time
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

'Image Stuff
Private Declare Function ImageList_Draw Lib "comctl32" (ByVal himl As Long, ByVal i As Long, ByVal hDCDest As Long, ByVal x As Long, ByVal Y As Long, ByVal flags As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Boolean

'Text Size
Private Declare Function GetTextExtentPoint32 Lib "gdi32" (ByVal hdc As Long, ByVal lpString As String, ByVal cbString As Long, lpSize As TextSize) As Boolean

'Memory stuff
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)

Public FileSearchCount As Long
Public FilesFound As Long
Public RecurseAmmount As Long
Public CurrentName As String
Public Abort As Boolean

Private Options_DisplayFullName As Boolean
Private Options_DisplayFiles As Boolean
Private Options_DisplayFolders As Boolean
Private Options_MinSize As Long
Private Options_MaxSize As Long
Private Options_DisplayHidden As Boolean
Private Options_DisplayArchive As Boolean
Private Options_DisplayReadOnly As Boolean
Private Options_DisplaySystem As Boolean

Private CURWFD As WIN32_FIND_DATA

Function FileGetNext(Path As String, hSearch As Long, Data As tFile) As Long
    FileGetNext = FindNextFile(hSearch, CURWFD)
    DataToFile Path, CURWFD, Data
End Function
Sub DataToFile(Path As String, WFD As WIN32_FIND_DATA, Data As tFile)
    With Data
        'Strings need to be converted
        .Name = StripNulls(WFD.FileName)
        .Path = Path
        .Size = (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
        .Attr = 0
        If WFD.dwFileAttributes And ATTR_ARCHIVE Then .Attr = .Attr Or vbArchive
        If WFD.dwFileAttributes And ATTR_DIRECTORY Then .Attr = .Attr Or vbDirectory
        If WFD.dwFileAttributes And ATTR_HIDDEN Then .Attr = .Attr Or vbHidden
        If WFD.dwFileAttributes And ATTR_NORMAL Then .Attr = .Attr Or vbNormal
        If WFD.dwFileAttributes And ATTR_READONLY Then .Attr = .Attr Or vbReadOnly
        If WFD.dwFileAttributes And ATTR_SYSTEM Then .Attr = .Attr Or vbSystem
    End With
End Sub
Private Function StripNulls(Str As String) As String
    Dim POS As Long
    POS = InStr(1, Str, vbNullChar)
    If POS Then StripNulls = Left$(Str, POS - 1) Else StripNulls = Str
End Function
Sub AddItem(TheListbox As ListBox, TheText As String)
    On Error Resume Next
    
    Call SendMessageAny(TheListbox.hwnd, LB_ADDSTRING, 0, ByVal TheText)
    
    Dim TextWidth As Long
    TextWidth = TheListbox.Parent.TextWidth(TheText) + 10
    If TextWidth > CLng(TheListbox.Tag) Then
        TheListbox.Tag = TextWidth
        Call AddHorizontalScrollBar(TheListbox, TextWidth)
    End If
End Sub
Function AddHorizontalScrollBar(TheListbox As ListBox, Pixels As Long) As Long
    AddHorizontalScrollBar = SendMessage(TheListbox.hwnd, LB_SETHORIZONTALEXTENT, Pixels, 0&)
End Function


Function GetRecurseFoldersListBox(TheListbox As ListBox, ByVal Directory As String, Filter As String, Count As Long, Files() As tFile) As Long
    Dim File As tFile, StartCount As Long, i As Long, hSearch As Long
    StartCount = Count
    
    hSearch = FindFirstFile(Directory & "*", CURWFD)
    If hSearch = INVALID_HANDLE_VALUE Then Exit Function

    Do
        If File.Name <> "." And File.Name <> ".." And File.Name <> vbNullString Then
            DoEvents    'Translate messages
            If Count > UBound(Files) Then ReDim Preserve Files(Count + ArrGrow)
            With Files(Count)
                .Path = Directory
                .Attr = File.Attr
                If .Attr And vbDirectory Then
                    .Name = File.Name & "\"
                    CurrentName = .Path & .Name
                    .FullName = CurrentName
                Else
                    .Name = File.Name
                    .Size = File.Size
                    .FullName = File.Path & File.Name
                End If
            End With
            
            Count = Count + 1
            FileSearchCount = FileSearchCount + 1
        End If
    Loop While FileGetNext(Directory, hSearch, File) <> 0 And (Abort = False)
    FindClose hSearch
    
    'IF THE FILE IS A DIRECTORY THEN ONLY DISPLAY THE FILE IF SHOWDIRECTORY = TRUE
    'IF THE FILE IS A FILE THEN ONLY DISPLAY THE FILE IF SHOWFILE = TRUE
    'IF THE FILE.HIDDEN THEN ONLY DISPLAY THE FILE IF SHOWHIDDEN = TRUE
    'IF THE FILE.READONLY THEN ONLY DISPLAY THE FILE IF SHOWREADONLY = TRUE
    'IF THE FILE.ARCHIVE THEN ONLY DISPLAY THE FILE IF SHOWARCHIVE = TRUE
    
    For i = StartCount To Count - 1
        If (Files(i).Size >= Options_MinSize Or Files(i).Size <= Options_MaxSize) And _
        ((Files(i).Attr And vbDirectory) = 0 Or Options_DisplayFiles) And _
        ((Files(i).Attr And vbDirectory) <> 0 Or Options_DisplayFolders) And _
        ((Files(i).Attr And vbReadOnly) <> 0 Or Options_DisplayReadOnly) And _
        ((Files(i).Attr And vbArchive) <> 0 Or Options_DisplayArchive) And _
        ((Files(i).Attr And vbHidden) <> 0 Or Options_DisplayHidden) And _
        ((Files(i).Attr And vbSystem) <> 0 Or Options_DisplaySystem) And _
        InStr(1, Files(i).Name, Filter, vbTextCompare) <> 0 Then
            Call AddItem(TheListbox, Files(i).FullName)
            FilesFound = FilesFound + 1
        End If
        If Files(i).Attr And vbDirectory Then GetRecurseFoldersListBox TheListbox, Files(i).FullName, Filter, Count, Files
NextItem:
    Next
End Function
Private Sub SearchStart(Files() As tFile)
    ReDim Files(ArrGrow)
    Abort = False
    FileSearchCount = 0
    FilesFound = 0
End Sub

Function FileSearch(ListBox As ListBox, Directory As String, Filter As String, Optional MinSize As Long = 0, Optional MaxSize As Long = -1, _
Optional ShowFiles As Boolean = True, Optional ShowFolders As Boolean = True, _
Optional ShowReadOnly As Boolean = True, Optional ShowArchive As Boolean = True, Optional ShowHidden As Boolean = True, _
Optional ShowSystem As Boolean = True _
) As tFile()
    'Our variables
    Dim Files() As tFile
    Dim Count As Long
    
    'Start the search
    Call SearchStart(Files)
    
    'Clear the list box
    ListBox.Clear
    
    'Make sure the Directory is right
    If Right(Directory, 1) <> "\" Then Directory = Directory & "\"
    
    'Set the module level variables for no OUT OF STACK SPACE ERRORS
    Options_MinSize = MinSize
    If MaxSize = -1 Then Options_MaxSize = MaxLong Else Options_MaxSize = MaxSize
    Options_DisplayFiles = Not ShowFiles
    Options_DisplayFolders = Not ShowFolders
    Options_DisplayReadOnly = Not ShowReadOnly
    Options_DisplayHidden = Not ShowHidden
    Options_DisplayArchive = Not ShowArchive
    Options_DisplaySystem = Not ShowSystem
    
    'Recursivly get folders and files
    Call GetRecurseFoldersListBox(ListBox, Directory, Filter, Count, Files)
    
    'Resize the files to only how much we found, remove the padding
    On Error Resume Next
    ReDim Preserve Files(0 To Count - 1)
    
    'Return the files we found
    FileSearch = Files
End Function


















