Attribute VB_Name = "Module2"
Option Explicit


Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Public Enum DRIVE_TYPE
    DRIVE_DOESNT_EXIST = 1
    DRIVE_REMOVABLE = 2
    DRIVE_FIXED = 3
    DRIVE_REMOTE = 4
    DRIVE_CDROM = 5
    DRIVE_RAMDISK = 6
End Enum
Private plLastDllError As Long

Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
'get win dir
Public Declare Function GetWindowsDirectory Lib "kernel32" _
Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, _
ByVal nSize As Long) As Long


Public Declare Function DllGetVersion _
    Lib "Shlwapi.dll" _
    (dwVersion As DllVersionInfo) As Long
    
 
'system startup
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
  
    
'clear recent documents
Public Declare Sub SHAddToRecentDocs Lib "shell32.dll" _
 (ByVal uFlags As Long, ByVal PV As String)
 
'display properties
Public Enum DISPLAY_PROPERTIES
    DP_Background
    DP_ScreenSaver
    DP_Appearance
    DP_Settings
End Enum



'ie version installed

 
Public Type DllVersionInfo
    cbSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
End Type



    
    


 
 'make ie trans
 Public _
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal _
dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal _
hWnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags _
As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
(ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias _
"FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 _
As String, ByVal lpsz2 As String) As Long
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Const WS_EX_TRANSPARENT = &H20&
Public Const LWA_ALPHA = &H2&



'taskbar

'public Declare Function FindWindow Lib "user32" Alias _
'"FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName _
'As String) As Long

Public Declare Function SetWindowPos Lib "user32" _
(ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_SHOWWINDOW = &H40


'cpu speed
  Public Const sCPURegKey = "HARDWARE\DESCRIPTION\System\CentralProcessor\0"

  Public Const HKEY_LOCAL_MACHINE As Long = &H80000002

  Public Declare Function RegCloseKey Lib "advapi32.dll" _
   (ByVal hKey As Long) As Long

  Public Declare Function RegOpenKey Lib "advapi32.dll" _
   Alias "RegOpenKeyA" _
  (ByVal hKey As Long, _
   ByVal lpSubKey As String, _
   phkResult As Long) As Long

  Public Declare Function RegQueryValueEx Lib "advapi32.dll" _
   Alias "RegQueryValueExA" _
   (ByVal hKey As Long, _
   ByVal lpValueName As String, _
   ByVal lpReserved As Long, _
   lpType As Long, _
   lpData As Any, _
   lpcbData As Long) As Long
'recyclebin
Public Declare Function ShellExecute Lib "shell32.dll" _
  Alias "ShellExecuteA" (ByVal hWnd As Long, _
  ByVal lpOperation As String, _
  ByVal lpFile As String, _
  ByVal lpParameters As String, _
  ByVal lpDirectory As String, _
  ByVal nShowCmd As Long) _
  As Long
  
Public Const SW_SHOWNORMAL As Long = 1

'startbtn
Public Declare Function ShowWindow Lib "user32" _
    (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    

    
'public Declare Function FindWindowEx Lib "user32" _
    Alias "FindWindowExA" (ByVal hWnd1 As Long, _
    ByVal hWnd2 As Long, _
    ByVal lpsz1 As String, _
    ByVal lpsz2 As String) As Long
    
    
    


'sreen saver
Public Declare Function SendMessage Lib "user32" Alias _
   "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
   ByVal wParam As Long, ByVal lParam As Long) As Long
Const WM_SYSCOMMAND = &H112&
Const SC_SCREENSAVE = &HF140&



Public Const SEE_MASK_INVOKEIDLIST As Long = &HC

Public Const SEE_MASK_NOCLOSEPROCESS As Long = &H40

Public Const SEE_MASK_FLAG_NO_UI As Long = &H400

Public Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Public Declare Function ShellExecuteEx Lib "shell32.dll" (SEI As SHELLEXECUTEINFO) As Long

'show file properties
Public Sub ShowFileProperties(ByVal Filename As String, ByVal OwnerhWnd As Long)

    On Error Resume Next
    Dim SEI As SHELLEXECUTEINFO
    Dim R As Long
    With SEI
        'Set the structure's size
        .cbSize = Len(SEI)
        'Seet the mask
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        'Set the owner window
        .hWnd = OwnerhWnd
        'Show the properties
        .lpVerb = "properties"
        'Set the filename
        .lpFile = Filename
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = 0
        .lpIDList = 0
    End With
    R = ShellExecuteEx(SEI)
End Sub

Public Function File_Open(Filename As String, Action As String) As Long
'
' Opens any file in its associated program
'

On Error Resume Next
    Dim Scr_hDC As Long
    Scr_hDC = GetDesktopWindow()
    File_Open = ShellExecute(Scr_hDC, Action, Filename, "", Left(Filename, 3), 1)
    If File_Open = 31 Then MsgBox "Failed to open this file, the associated program might not exist !", vbCritical, "Crap..."
Exit Function
End Function




Private Function PathCheck(ByVal PathName As String, Optional AltDelimiter As String = "") As String
    If Len(PathName) = 0 Then Exit Function
    Dim Delimiter As String
    Delimiter = IIf(InStr(PathName, "/"), "/", "\")
    PathCheck = IIf(Right$(PathName, 1) = Delimiter, PathName, PathName & Delimiter)
    PathCheck = IIf(Len(AltDelimiter) = 0, PathCheck, Replace(PathCheck, Delimiter, AltDelimiter))
End Function

Public Function DriveMBSize(Optional Drive As String = "C:\") As Double
    'some time in the future disk may be to large to calculate
    'like this so resume next on any errors
    On Error Resume Next

    Dim lAns As Long
    Dim lSectorsPerCluster As Long
    Dim lBytesPerSector As Long
    
    Dim lFreeClusters As Long
    Dim lTotalClusters As Long
    Dim lBytesPerCluster As Long
    Dim lTotalBytes As Double
    
    'fix bad parameter values
    If Len(Drive) = 1 Then Drive = Drive & ":\"
    If Len(Drive) = 2 And Right$(Drive, 1) = ":" Then Drive = Drive & "\"
    
    lAns = GetDiskFreeSpace(Drive, lSectorsPerCluster, lBytesPerSector, lFreeClusters, lTotalClusters)
    lBytesPerCluster = lSectorsPerCluster * lBytesPerSector
    
    DriveMBSize = ((lBytesPerCluster / 1024) / 1024) * lTotalClusters
    DriveMBSize = Format(DriveMBSize, "###,###,##0.00")
End Function

Public Function DriveMBFree(Optional Drive As String = "C:\") As Double
    'some time in the future disk may be to large to calculate
    'like this so resume next on any errors
    On Error Resume Next

    Dim lAns As Long
    Dim lSectorsPerCluster As Long
    Dim lBytesPerSector As Long

    Dim lFreeClusters As Long
    Dim lTotalClusters As Long
    Dim lBytesPerCluster As Long
    Dim lFreeBytes As Double

    'fix bad parameter values
    If Len(Drive) = 1 Then Drive = Drive & ":\"
    If Len(Drive) = 2 And Right$(Drive, 1) = ":" Then Drive = Drive & "\"

    lAns = GetDiskFreeSpace(Drive, lSectorsPerCluster, lBytesPerSector, lFreeClusters, lTotalClusters)
    lBytesPerCluster = lSectorsPerCluster * lBytesPerSector

    DriveMBFree = ((lBytesPerCluster / 1024) / 1024) * lFreeClusters
    DriveMBFree = Format(DriveMBFree, "###,###,##0.00")
End Function

Public Function DriveType(Drive As String) As DRIVE_TYPE
    'fix bad parameter values
    DriveType = DRIVE_DOESNT_EXIST
    plLastDllError = 0
    Drive = IIf(Len(Drive) = 1, Drive & ":", Drive)
    If Len(Drive) = 1 Then Drive = Drive & ":\"
    If Len(Drive) = 2 And Right$(Drive, 1) = ":" Then Drive = Drive & "\"
    Drive = PathCheck(Drive)
    DriveType = GetDriveType(Drive)
    plLastDllError = Err.LastDllError
End Function


