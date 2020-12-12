VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
End
Attribute VB_Name = "CShellExecute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"No"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

' Class       : CShellExecute
' Description : Use the Windows Shell to launch, print,
'               or explore files or programs
' Source      : Total VB SourceBook 6

' Declarations for Windows API calls
Private Declare Function ShellExecute _
  Lib "shell32.dll" _
  Alias "ShellExecuteA" _
  (ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) _
  As Long

Private Declare Sub ShellExecuteEx _
  Lib "shell32.dll" _
  Alias "ShellExecuteExA" _
  (lpExecInfo As SHELLEXECUTEINFO)

Private Declare Function GetExitCodeProcess _
  Lib "kernel32" _
  (ByVal hProcess As Long, _
  lpExitCode As Long) _
  As Long
  
Private Declare Function OpenProcess _
  Lib "kernel32" _
  (ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) _
  As Long

Private Type SHELLEXECUTEINFO
  cbSize As Long
  fMask As Long
  hwnd As Long
  lpVerb As String
  lpFile As String
  lpParameters As String
  lpDirectory As String
  nShow As Long
  hInstApp As Long
  '  Optional fields
  lpIDList As Long
  lpClass As String
  hkeyClass As Long
  dwHotKey As Long
  hIcon As Long
  hProcess As Long
End Type

Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_DOENVSUBST = &H200

' Public property enumerated constants
Public Enum EnumShellExecuteErrors
  seeNoError = -1                 'Any value above 32
  seeOUT_OF_MEMORY = 0            'The operating system is out of memory or resources.
  seeERROR_FILE_NOT_FOUND = 2     'The specified file was not found.
  seeERROR_PATH_NOT_FOUND = 3     'The specified path was not found.
  seeERROR_BAD_FORMAT = 11        'The .exe file is invalid (non-Win32® .exe or error in .exe image).
  seeSE_ERR_ACCESSDENIED = 5      'The operating system denied access to the specified file.
  seeSE_ERR_ASSOCINCOMPLETE = 27  'The file name association is incomplete or invalid.
  seeSE_ERR_DDEBUSY = 30          'The DDE transaction could not be completed because other DDE transactions were being processed.
  seeSE_ERR_DDEFAIL = 29          'The DDE transaction failed.
  seeSE_ERR_DDETIMEOUT = 28       'The DDE transaction could not be completed because the request timed out.
  seeSE_ERR_DLLNOTFOUND = 32      'The specified dynamic-link library was not found.
  seeSE_ERR_NOASSOC = 31          'There is no application associated with the given file name extension.
  seeSE_ERR_OOM = 8               'There was not enough memory to complete the operation.
  seeSE_ERR_SHARE = 26            'A sharing violation occurred.

End Enum

Public Enum EnumShellExecuteShowStyles
  sesSW_HIDE = 0
  sesSW_MAXIMIZE = 3
  sesSW_MINIMIZE = 6
  sesSW_RESTORE = 9
  sesSW_SHOW = 5
  sesSW_SHOWDEFAULT = 10
  sesSW_SHOWMAXIMIZED = 3
  sesSW_SHOWMINIMIZED = 2
  sesSW_SHOWMINNOACTIVE = 7
  sesSW_SHOWNA = 8
  sesSW_SHOWNOACTIVATE = 4
  sesSW_SHOWNORMAL = 1
End Enum

Private mlnghInstance As Long
Private mlnghProcess As Long

Public Property Get StillExecuting() As Boolean
  ' Returns: A true or false value indicating whether the application
  '          launched with the LaunchProgram or LaunchDocument method
  '          is still executing
  ' Source: Total VB SourceBook 6
  '
  Const clngStillActive = &H103
    
  Dim lngRet As Long
  Dim fResult As Boolean
  Dim lngExit As Long
  
  On Error GoTo PROC_ERR
  
  ' This variable is set if an app is successfully launched
  ' with either LaunchProgram or LaunchDocument
  
  If mlnghInstance <> 0 Then
    ' Test whether or not the app is still running
    lngRet = GetExitCodeProcess(mlnghProcess, lngExit)
    
    If lngRet > 0 Then
      fResult = (lngExit = clngStillActive)
    Else
      fResult = False
    End If
      
  Else
    fResult = False
  End If
  
  StillExecuting = fResult
  
PROC_EXIT:
  Exit Property

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "StillExecuting"
  Resume PROC_EXIT

End Property

Public Function Explore( _
  lnghWnd As Long, _
  strFolder As String, _
  Optional eShowStyle As EnumShellExecuteShowStyles = sesSW_SHOWDEFAULT) _
  As EnumShellExecuteErrors
  ' Comments  : Launches Windows Explorer open to the
  '             specified folder
  ' Parameters: lnghWnd - Handle to window of a form
  '             strFolder - Path to the directory to explore
  '             eShowStyle - Constant indicating how the folder
  '             is to be displayed (maximized, minimized etc.)
  ' Returns   : -1 on Success, or one of the values in the
  '             EnumShellExecuteErrors constants on failure
  ' Source    : Total VB SourceBook 6
  '
  Dim lngResult As Long
  
  On Error GoTo PROC_ERR
  
  mlnghInstance = 0
  mlnghProcess = 0
  
  lngResult = ShellExecute( _
    lnghWnd, _
    "explore", _
    strFolder, _
    vbNullString, _
    vbNullString, _
    eShowStyle)
    
  If lngResult > 32 Then
    Explore = seeNoError
  Else
    Explore = lngResult
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "Explore"
  Resume PROC_EXIT
  
End Function

Public Function LaunchDocument( _
  lnghWnd As Long, _
  strDocument As String, _
  strStartIn As String, _
  Optional eShowStyle As EnumShellExecuteShowStyles = sesSW_SHOWDEFAULT) _
  As EnumShellExecuteErrors
  ' Comments  : Launch a program by specifying the document
  '             or data file. For example this could be a Word
  '             document or an Access .MDB database
  ' Parameters: lnghWnd - Handle to window of a form
  '             strDocument - Path to the document to launch
  '             strStartIn - Default startup directory for
  '             the program
  '             eShowStyle - Constant indicating how the program
  '             is to be displayed (maximized, minimized etc.)
  ' Returns   : -1 on Success, or one of the values in the
  '             EnumShellExecuteErrors constants on failure
  ' Source    : Total VB SourceBook 6
  '
  Dim lngResult As Long
  Dim sexinfo As SHELLEXECUTEINFO
  
  On Error GoTo PROC_ERR
  
  mlnghInstance = 0
  mlnghProcess = 0
  
  With sexinfo
    .cbSize = Len(sexinfo)
    .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_DOENVSUBST
    .hwnd = lnghWnd
    .lpVerb = "Open"
    .lpFile = strDocument
    .lpParameters = ""
    .lpDirectory = strStartIn
    .nShow = eShowStyle
  End With
  
  ShellExecuteEx sexinfo
  
  lngResult = sexinfo.hInstApp
  
  If lngResult > 32 Then
    LaunchDocument = seeNoError
  Else
    LaunchDocument = lngResult
  End If
  
  mlnghInstance = lngResult
  mlnghProcess = sexinfo.hProcess
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "LaunchDocument"
  Resume PROC_EXIT
End Function

Public Function LaunchEmail( _
  lnghWnd As Long, _
  strAddress As String, _
  Optional eShowStyle As EnumShellExecuteShowStyles = sesSW_SHOWDEFAULT) _
  As EnumShellExecuteErrors
  ' Comments  : Open the program associated with email on the system.
  ' Parameters: lnghWnd - Handle to window of a form
  '             strAddress - email address in the form user@domain.com
  '             eShowStyle - Constant indicating how the browser
  '             is to be displayed (maximized, minimized etc.)
  ' Returns   : -1 on Success, or one of the values in the
  '             EnumShellExecuteErrors constants on failure
  ' Source    : Total VB SourceBook 6
  '
  Dim lngResult As Long
  Dim strTmp As String
  
  On Error GoTo PROC_ERR
  
  mlnghInstance = 0
  mlnghProcess = 0
  
  strTmp = "mailto:" & strAddress
  
  lngResult = ShellExecute( _
    lnghWnd, _
    vbNullString, _
    strTmp, _
    vbNullString, _
    vbNullString, _
    eShowStyle)
    
  If lngResult > 32 Then
    LaunchEmail = seeNoError
  Else
    LaunchEmail = lngResult
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "LaunchEmail"
  Resume PROC_EXIT
  
End Function

Public Function LaunchFolder( _
  lnghWnd As Long, _
  strFolder As String, _
  Optional eShowStyle As EnumShellExecuteShowStyles = sesSW_SHOWDEFAULT) _
  As EnumShellExecuteErrors
  ' Comments  : Open a folder window on a directory
  ' Parameters: lnghWnd - Handle to window of a form
  '             strFolder - Path to the directory to open
  '             eShowStyle - Constant indicating how the folder
  '             is to be displayed (maximized, minimized etc.)
  ' Returns   : -1 on Success, or one of the values in the
  '             EnumShellExecuteErrors constants on failure
  ' Source    : Total VB SourceBook 6
  '
  Dim lngResult As Long
  
  On Error GoTo PROC_ERR
  
  mlnghInstance = 0
  mlnghProcess = 0
  
  lngResult = ShellExecute( _
    lnghWnd, _
    "open", _
    strFolder, _
    vbNullString, _
    vbNullString, _
    eShowStyle)
    
  If lngResult > 32 Then
    LaunchFolder = seeNoError
  Else
    LaunchFolder = lngResult
  End If

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "LaunchFolder"
  Resume PROC_EXIT

End Function

Public Function LaunchProgram( _
  lnghWnd As Long, _
  strProgram As String, _
  Optional strParameters As String = vbNullString, _
  Optional strStartIn As String = vbNullString, _
  Optional eShowStyle As EnumShellExecuteShowStyles = sesSW_SHOWDEFAULT) _
  As EnumShellExecuteErrors
  ' Comments  : Launch a program
  ' Parameters: lnghWnd - Handle to window of a form
  '             strProgram - Path to program to launch
  '             strParameters - Command line arguments for
  '             the program
  '             strStartIn - Default startup directory for
  '             the program
  '             eShowStyle - Constant indicating how the program
  '             is to be displayed (maximized, minimized etc.)
  ' Returns   : -1 on Success, or one of the values in the
  '             EnumShellExecuteErrors constants on failure
  ' Source    : Total VB SourceBook 6
  '
  Dim lngResult As Long
  Dim sexinfo As SHELLEXECUTEINFO
  
  On Error GoTo PROC_ERR
  
  mlnghInstance = 0
  mlnghProcess = 0
  
  With sexinfo
    .cbSize = Len(sexinfo)
    .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_DOENVSUBST
    .hwnd = lnghWnd
    .lpVerb = "Open"
    .lpFile = strProgram
    .lpParameters = strParameters
    .lpDirectory = strStartIn
    .nShow = eShowStyle
  End With
  
  ShellExecuteEx sexinfo
  
  lngResult = sexinfo.hInstApp
  
  If lngResult > 32 Then
    LaunchProgram = seeNoError
  Else
    LaunchProgram = lngResult
  End If
  
  mlnghInstance = lngResult
  mlnghProcess = sexinfo.hProcess

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "LaunchProgram"
  Resume PROC_EXIT

End Function

Public Function LaunchURL( _
  lnghWnd As Long, _
  strURL As String, _
  Optional eShowStyle As EnumShellExecuteShowStyles = sesSW_SHOWDEFAULT) _
  As EnumShellExecuteErrors
  ' Comments  : Open the default Internet browser on the user's machine
  '             by specifying a URL (uniform resource locator), for example:
  '             http://www.fmsinc.com
  ' Parameters: lnghWnd - Handle to window of a form
  '             strURL - URL to open
  '             eShowStyle - Constant indicating how the browser
  '             is to be displayed (maximized, minimized etc.)
  ' Returns   : -1 on Success, or one of the values in the
  '             EnumShellExecuteErrors constants on failure
  ' Source    : Total VB SourceBook 6
  '
  Dim lngResult As Long
  
  On Error GoTo PROC_ERR
  
  mlnghInstance = 0
  mlnghProcess = 0
    
  lngResult = ShellExecute( _
    lnghWnd, _
    vbNullString, _
    strURL, _
    vbNullString, _
    vbNullString, _
    eShowStyle)
    
  If lngResult > 32 Then
    LaunchURL = seeNoError
  Else
    LaunchURL = lngResult
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "LaunchURL"
  Resume PROC_EXIT
End Function

Public Function PrintDocument( _
  lnghWnd As Long, _
  strDocument As String) _
  As EnumShellExecuteErrors
  ' Comments  : Prints a document by launching the program
  '             that is associated with the document (similar
  '             to what happens when you drag a document to the
  '             printer icon on the desktop.)
  ' Parameters: lnghWnd - Handle to window of a form
  '             strDocument - Path to the document to print
  ' Returns   : -1 on Success, or one of the values in the
  '             EnumShellExecuteErrors constants on failure
  ' Source    : Total VB SourceBook 6
  '
  Dim lngResult As Long
  
  On Error GoTo PROC_ERR
  
  mlnghInstance = 0
  mlnghProcess = 0
  
  lngResult = ShellExecute( _
    lnghWnd, _
    "print", _
    strDocument, _
    vbNullString, _
    vbNullString, _
    0)
    
  If lngResult > 32 Then
    PrintDocument = seeNoError
  Else
    PrintDocument = lngResult
  End If

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "PrintDocument"
  Resume PROC_EXIT

End Function

