Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function auxSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Public Declare Function auxGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Public Declare Function CloseWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As Long)
Public Declare Function CreateSolidBrush Lib "GDI32" (ByVal crColor As Long) As Long
Public Declare Function CreateWindow Lib "user32" Alias "CreateWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function CreateWindowEx Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Public Declare Function EnumDesktopWindows Lib "user32" (ByVal hDesktop As Long, ByVal lpfn As Long, ByVal lParam As Long) As Boolean
Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function fCreateShellLink Lib "STKIT432.DLL" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Public Declare Function GetCapture Lib "user32" () As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function GetHostName Lib "WSOCK32.DLL" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Public Declare Function GetHostByName Lib "WSOCK32.DLL" (ByVal szHost As String) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Integer
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hWndCallback As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function Rectangle Lib "GDI32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SetBkColor Lib "GDI32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "user32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SHAddToRecentDocs Lib "Shell32" (ByVal lFlags As Long, ByVal lPv As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "Shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function StretchBlt Lib "GDI32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Public Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

'Window Messages
Public Const WM_CHAR = &H102
Public Const WM_SETTEXT = &HC
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_DESTROY = &H2
Public Const WM_MOVE = &HF012
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_SYSCOMMAND = &H112
Public Const WM_MOUSEMOVE = &H200
Public Const WM_CLEAR = &H303
Public Const WM_DRAWITEM = &H2B
Public Const WM_PAINT = &HF
Public Const WM_ERASEBKGND = &H14
Public Const WM_NCPAINT = &H85



'Combo Box Functions
Public Const CB_DELETESTRING = &H144
Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_RESETCONTENT = &H14B
Public Const CB_SETCURSEL = &H14E
Public Const CB_GETCOUNT = &H146

'hWnd Functions
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

'Show Window Functions
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
Public Const SW_NORMAL = 1

'Sound Functions
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT
Public Const SND_LOOP = &H8

'Screen Saver Function
Public Const SPI_SCREENSAVERRUNNING = 97

'Get Window Word Functions
Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)

'Virtual Key Statements
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26

'Phader Color Presets
Public Const COLOR_RED = &HFF&
Public Const COLOR_GREEN = &HFF00&
Public Const COLOR_BLUE = &HFF0000
Public Const COLOR_YELLOW = &HFFFF&
Public Const COLOR_WHITE = &HFFFFFE
Public Const COLOR_BLACK = &H0&
Public Const COLOR_PEACH = &HC0C0FF
Public Const COLOR_PURPLE = &HFF00FF
Public Const COLOR_GREY = &HC0C0C0
Public Const COLOR_PINK = &HFF80FF
Public Const COLOR_TURQUOISE = &HC0C000
Public Const COLOR_LIGHTBLUE = &HFF8080
Public Const COLOR_ORANGE = &H80FF&

'Processor Types
Public Const PROCESSOR_INTEL_386 = 386
Public Const PROCESSOR_INTEL_486 = 486
Public Const PROCESSOR_INTEL_PENTIUM = 586
Public Const PROCESSOR_MIPS_R4000 = 4000
Public Const PROCESSOR_ALPHA_21064 = 21064

'Menu Functions
Public Const MF_DELETE = &H200&
Public Const MF_CHANGE = &H80&
Public Const MF_ENABLED = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_POPUP = &H10&
Public Const MF_UNCHECKED = &H0&
Public Const MF_CHECKED = &H8&

'Key Presets
Public Const ENTER_KEY = 13

'Button Messages
Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1

'List Box Functions
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETCOUNT = &H18B
Public Const LB_ADDSTRING = &H180
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_GETCURSEL = &H188
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SELECTSTRING = &H18C
Public Const LB_SETCOUNT = &H1A7
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const LB_INSERTSTRING = &H181
Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000

'Notify Icon Functions
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIM_MODIFY = &H1
Public Const NIF_TIP = &H4

'Edit Window Messages
Public Const EM_REPLACESEL = &HC2
Public Const EM_SETSEL = &HB1

'Dev Mode Const's
Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000

'Windows Version Functions
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

'Byte Functions
Public Const MAX_DEFAULTCHAR = 2
Public Const MAX_LEADBYTES = 12

'winsck functions
Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128
Public Const ERROR_SUCCESS       As Long = 0
Public Const WS_VERSION_REQD     As Long = &H101
Public Const WS_VERSION_MAJOR    As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR    As Long = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD    As Long = 1
Public Const SOCKET_ERROR        As Long = -1

'types
Public Type HOSTENT
    hName      As Long
    hAliases   As Long
    hAddrType  As Integer
    hLen       As Integer
    hAddrList  As Long
End Type

Public Type WSADATA
    wVersion      As Integer
    wHighVersion  As Integer
    szDescription(0 To MAX_WSADescription)   As Byte
    szSystemStatus(0 To MAX_WSASYSStatus)    As Byte
    wMaxSockets   As Integer
    wMaxUDPDG     As Integer
    dwVendorInfo  As Long
End Type

Public Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer

    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer

    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Dim DevM As DEVMODE

Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type SYSTEM_INFO
        dwOemID As Long
        dwPageSize As Long
        lpMinimumApplicationAddress As Long
        lpMaximumApplicationAddress As Long
        dwActiveProcessorMask As Long
        dwNumberOrfProcessors As Long
        dwProcessorType As Long
        dwAllocationGranularity As Long
        dwReserved As Long
End Type

Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
End Type

Public Type CPINFO
        MaxCharSize As Long
        DefaultChar(MAX_DEFAULTCHAR - 1) As Byte
        LeadByte(MAX_LEADBYTES - 1) As Byte
End Type

Public Type NOTIFYICONDATA
        cbSize As Long
        hWnd As Long
        uId As Long
        uFlags As Long
        ucallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Public Type COLORRGB
    Red As Long
    Green As Long
    Blue As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

'Enums
Public Enum aolshortcutkeys
    ctrl1
    ctrl2
    ctrl3
    ctrl4
    ctrl5
    ctrl6
    ctrl7
    ctrl8
    ctrl9
    ctrl0
End Enum
Function color1(color, frm As Form)
On Error Resume Next
Dim Crl As Control
For Each Crl In frm.Controls
    Crl.BackColor = color
Next Crl
frm.BackColor = color
End Function
Function SendIM(sn As String, message As String)
Dim oscarbuddylistwin&
Dim oscartabgroup&
Dim oscariconbtn&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
oscartabgroup& = FindWindowEx(oscarbuddylistwin&, 0&, "_oscar_tabgroup", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, 0&, "_oscar_iconbtn", vbNullString)

SendIM = oscariconbtn&
Call leftclick(SendIM)

Dim aimimessage&
Dim oscarpersistantcombo&
Dim edit&
aimimessage& = FindWindow("aim_imessage", vbNullString)
oscarpersistantcombo& = FindWindowEx(aimimessage&, 0&, "_oscar_persistantcombo", vbNullString)
edit& = FindWindowEx(oscarpersistantcombo&, 0&, "edit", vbNullString)

SendIM = edit&
Call SendMessageByString(SendIM, WM_SETTEXT, WM_CHAR, sn$)

Dim wndateclass&
Dim ateclass&
aimimessage& = FindWindow("aim_imessage", vbNullString)
wndateclass& = FindWindowEx(aimimessage&, 0&, "wndate32class", vbNullString)
wndateclass& = FindWindowEx(aimimessage&, wndateclass&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)

SendIM = ateclass&
Call SendMessageByString(SendIM, WM_SETTEXT, WM_CHAR, message$)


aimimessage& = FindWindow("aim_imessage", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, 0&, "_oscar_iconbtn", vbNullString)

SendIM = oscariconbtn&
Call leftclick(SendIM)
End Function
Function clearim()
Dim aimimessage&
Dim wndateclass&
Dim ateclass&
Dim clear&
aimimessage& = FindWindow("aim_imessage", vbNullString)
wndateclass& = FindWindowEx(aimimessage&, 0&, "wndate32class", vbNullString)
wndateclass& = FindWindowEx(aimimessage&, wndateclass&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)

clear& = ateclass&
Call SetText(clear&, "")
End Function
Sub leftclick(hWnd)
Call SendMessage(hWnd, WM_LBUTTONDOWN, 0, 0)
Call SendMessage(hWnd, ENTER_KEY, 0, 0)
Call SendMessage(hWnd, WM_LBUTTONUP, 0, 0)
End Sub
Function preferences()

Dim oscarbuddylistwin&
Dim oscariconbtn&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
oscariconbtn& = FindWindowEx(oscarbuddylistwin&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(oscarbuddylistwin&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(oscarbuddylistwin&, oscariconbtn&, "_oscar_iconbtn", vbNullString)

preferences = oscariconbtn&
Call leftclick(preferences)
End Function
Function Bot_Lamerizer(Nam As String)
'Ex: Call Bot_Lamerizer(text1)
'You gotta add a text box to enter
'who to lamerize

Dim X As Integer
Dim lcse As String
Dim letr As String
Dim dis As String
SendChat (Nam$ & " IS A LAMER!!! HERES WHY!")
For X = 1 To Len(Nam)
lcse$ = LCase(Nam)
letr$ = Mid(lcse$, X, 1)
If letr$ = " " Then Let dis$ = "The blank is what your mind draws": GoTo Dissem
If letr$ = "a" Then Let dis$ = "'a' is for the animals utters your mom touches": GoTo Dissem
If letr$ = "b" Then Let dis$ = "'b' is for the bull dyke you are": GoTo Dissem
If letr$ = "c" Then Let dis$ = "'c' is for the cows your mom molests with heated branders": GoTo Dissem
If letr$ = "d" Then Let dis$ = "'d' is for all the dogs ur mom feed from her mouth": GoTo Dissem
If letr$ = "e" Then Let dis$ = "'e' is for the etheopian your mom is": GoTo Dissem
If letr$ = "f" Then Let dis$ = "'f' is for the fries you make at African McDonalds": GoTo Dissem
If letr$ = "g" Then Let dis$ = "'g' is for the garden gnomes your mom tries to lick": GoTo Dissem
If letr$ = "h" Then Let dis$ = "'h' is for the hose that your mom could suck a golf ball through": GoTo Dissem
If letr$ = "i" Then Let dis$ = "'i' is for the idiot you are in 'Coloring bewteen the lines'": GoTo Dissem
If letr$ = "j" Then Let dis$ = "'j' is for all the JuJu Bees your mom screws": GoTo Dissem
If letr$ = "k" Then Let dis$ = "'k' is for all the klowns in the circus your mom has affairs with": GoTo Dissem
If letr$ = "l" Then Let dis$ = "'l' is for lickings you mom gets from elephants": GoTo Dissem
If letr$ = "m" Then Let dis$ = "'m' is for how many moms you had": GoTo Dissem
If letr$ = "n" Then Let dis$ = "'n' is for the nickopheliac your dad is": GoTo Dissem
If letr$ = "o" Then Let dis$ = "'o' is for the orphan your parents were in 'Annie'": GoTo Dissem
If letr$ = "p" Then Let dis$ = "'p' poopy your family eats for fun": GoTo Dissem
If letr$ = "q" Then Let dis$ = "'q' is for the quimby you are (Somebody who sucks farts out of a dead chicken)": GoTo Dissem
If letr$ = "r" Then Let dis$ = "'r' is for the rapings you gave ur family": GoTo Dissem
If letr$ = "s" Then Let dis$ = "'s' is for the shots your got for genital herpes": GoTo Dissem
If letr$ = "t" Then Let dis$ = "'t' is for the test tube baby you were": GoTo Dissem
If letr$ = "u" Then Let dis$ = "'u' is for the utters you pet": GoTo Dissem
If letr$ = "v" Then Let dis$ = "'v' is for the ... think of something ... ": GoTo Dissem
If letr$ = "w" Then Let dis$ = "'w' is for the heavy weight in ur moms herpe":  GoTo Dissem
If letr$ = "x" Then Let dis$ = "'x' is for the xtreme zits that u have on ur mother": GoTo Dissem
If letr$ = "y" Then Let dis$ = "'y' is for the years of agony you go through knowing ur moms a dog": GoTo Dissem
If letr$ = "z" Then Let dis$ = "'z' is for the zebras you like to touch":  GoTo Dissem

If letr$ = "1" Then Let dis$ = "'1' is for the number of that you had a g/f.. that u will ever have": GoTo Dissem
If letr$ = "2" Then Let dis$ = "'2' is for the number of 3 world countries your mom worked at for McDonalds": GoTo Dissem
If letr$ = "3" Then Let dis$ = "'3' is for the number of dads you have": GoTo Dissem
If letr$ = "4" Then Let dis$ = "'4' is for the number of Animal Sibilings you have":  GoTo Dissem
If letr$ = "5" Then Let dis$ = "'5' is for the number of the same sex what u had you made out with": GoTo Dissem
If letr$ = "6" Then Let dis$ = "'6' is for the number of times you found out u had a different mom": GoTo Dissem
If letr$ = "7" Then Let dis$ = "'7' is for the times that u that a tree was a person": GoTo Dissem
If letr$ = "8" Then Let dis$ = "'8' is for how many moms you had": GoTo Dissem
If letr$ = "9" Then Let dis$ = "'9' is for how many moms that u dont know that u have": GoTo Dissem
If letr$ = "0" Then Let dis$ = "'0' is for the times you had friends": GoTo Dissem

Dissem:
Call SendChat(dis$)

Pause (3)
Next X

End Function
Function buddychat(screennames As String)

Dim oscarbuddylistwin&
Dim oscartabgroup&
Dim oscariconbtn&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
oscartabgroup& = FindWindowEx(oscarbuddylistwin&, 0&, "_oscar_tabgroup", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, oscariconbtn&, "_oscar_iconbtn", vbNullString)

buddychat = oscariconbtn&
Call leftclick(buddychat)

Dim aimchatinvitesendwnd&
Dim edit&
aimchatinvitesendwnd& = FindWindow("aim_chatinvitesendwnd", vbNullString)
edit& = FindWindowEx(aimchatinvitesendwnd&, 0&, "edit", vbNullString)

buddychat = edit&
Call SendMessageByString(buddychat, WM_SETTEXT, WM_CHAR, screennames$)


aimchatinvitesendwnd& = FindWindow("aim_chatinvitesendwnd", vbNullString)
oscariconbtn& = FindWindowEx(aimchatinvitesendwnd&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatinvitesendwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatinvitesendwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)

buddychat = oscariconbtn&
Call leftclick(buddychat)
Call leftclick(buddychat)
End Function

Function changecap(caption As String)
Dim oscarbuddylistwin&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)

changecap = oscarbuddylistwin&
Call SendMessageByString(changecap, WM_SETTEXT, WM_CHAR, caption$)
End Function

Function PixelsToTwips_height(pxls)
    PixelsToTwips_height = pxls * Screen.TwipsPerPixelY
End Function


Function PixelsToTwips_width(pxls)
    PixelsToTwips_width = pxls * Screen.TwipsPerPixelX
End Function


Function dragit(frm As Form)
  ReleaseCapture
    SendMessage frm.hWnd, &HA1, 2, 0&
End Function

Public Function aimchatline() As String
    On Error Resume Next
    Dim lngchatwin As Long, lngatewin As Long
    Dim lnglong As Long, strstring As String
    Let lngchatwin& = FindWindow("aim_chatwnd", vbNullString)
    Let lngatewin& = FindWindowEx(lngchatwin&, 0&, "wndate32class", vbNullString)
    Let strstring$ = GetText(lngatewin&)
    Let lnglong& = InStrRev(strstring$, "<br>")
    Let aimchatline$ = Right$(strstring$, Len(strstring$) - lnglong& - 3&)
End Function
Public Function Keep(String1 As String, LettersToKeep)
'allows u to tell it a string and it will take out all the characters that u dont tell it to keep
    Dim String2 As String, i As Integer, Letter As String, InString As Integer
    For i = 1 To Len(String1$)
        Letter$ = Mid(String1$, i, 1)
        InString = InStr(LettersToKeep, Letter$)
        If InString <> 0 Then
            String2$ = String2$ & Letter$
        End If
    Next i
    Keep = String2$
    
End Function
Public Function RemoveHTML(String1 As String)
'removes all the html characters form a string
    Dim FH As String, LH As String, LocOfLT As Long, LocOfGT As Long
    Do Until InStr(String1$, "<") = 0 Or InStr(String1$, ">") = 0
        If InStr(String1$, "<") > InStr(String1$, ">") Then
            Exit Do
        End If
        LocOfLT& = InStr(String1$, "<")
        LocOfGT& = InStr(String1$, ">")
        FH$ = Left$(String1$, LocOfLT& - 1)
        LH$ = Mid$(String1$, LocOfGT& + 1)
        String1$ = FH$ & LH$
    Loop
    RemoveHTML = String1$

End Function
Public Function Aim_GetChatText()
'Gets all the text from the aim 2.1+ chat textbox
    Dim Window As Long, Window1 As Long, ChatTB As Long, ChatTBLength As Long, buffer As String
    Window& = FindWindow("AIM_ChatWnd", vbNullString)
    Window1& = FindWindowEx(Window&, 0&, "WndAte32Class", "AteWindow")
    ChatTB& = FindWindowEx(Window1&, 0&, "Ate32Class", vbNullString)
    ChatTBLength& = SendMessage(ChatTB&, WM_GETTEXTLENGTH, 0&, 0&)
    buffer$ = String(ChatTBLength&, 0&)
    Call SendMessageByString(ChatTB&, WM_GETTEXT, ChatTBLength& + 1, buffer$)
    Aim_GetChatText = buffer

End Function
Public Function Aim_LastLine() As String
'my pride and joy...i wrote ALL this code it was not stolen
    On Error GoTo ErrHandler

    Dim ChatText As String
    ChatText$ = Aim_GetChatText
    If Len(ChatText$) > 500 Then
        ChatText$ = Right$(ChatText$, 250)
    End If
    If InStr(ChatText$, ")--></B></FONT><FONT COLOR=""#") <> 0 Then
        ChatText$ = Mid$(ChatText$, LastInStr(ChatText$, ")--></B></FONT><FONT COLOR=""#") + 38)
        ChatText$ = RemoveHTML(ChatText$)
        ChatText$ = Trim$(ChatText$)
        Aim_LastLine = Keep(ChatText$, "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890)(*&^%$#@!~`"";:'?\/.,][}{+=_-|» ")
        Aim_LastLine = Left$(Aim_LastLine, InStr(Aim_LastLine, "/") - 1&)
    Else
        Aim_LastLine = ""
    End If
    Exit Function
    
ErrHandler:
    Aim_LastLine = ""

End Function
Public Function Aim_LastSender() As String
'another one of my good functions...
    On Error GoTo ErrHandler
    
    Dim ChatText As String
    ChatText$ = Aim_GetChatText()
    If InStr(ChatText$, "<BODY BGCOLOR=""#") <> 0 Then
        If Len(ChatText$) > 500 Then
            ChatText$ = Right$(ChatText$, 250)
        End If
        ChatText$ = Mid$(ChatText$, LastInStr(ChatText$, "<BODY BGCOLOR=""#"))
        If InStr(ChatText$, "<!-- (") <> 0 Then
            ChatText$ = Left$(ChatText$, InStr(ChatText$, "<!-- (") - 1)
            ChatText$ = Mid$(ChatText$, LastInStr(ChatText$, ">") + 1)
            Aim_LastSender = ChatText$
        Else
            Aim_LastSender = ""
        End If
    Else
        Aim_LastSender = ""
    End If
    Exit Function
    
ErrHandler:
    Aim_LastSender = ""

End Function
Public Function LastInStr(String1 As String, WhatToFind As String)
'finds the last occurence of a string within a nother string
    Dim CurrLoc As Long, i As Long
    For i = 1 To Len(String1$) - Len(WhatToFind$) + 1

        If Mid$(String1$, i, Len(WhatToFind$)) = WhatToFind$ Then CurrLoc& = i
    
    Next i

    LastInStr = CurrLoc&

End Function
Public Function aimchatlinemsg() As String
    Dim strtext As String, lnglong As Long, lngatewin As Long
    Let strtext$ = aimchatline$
    Let lnglong& = InStrRev(strtext$, "<br>")
    Let strtext$ = Right$(strtext$, Len(strtext$) - lnglong& - 3&)
    Let strtext$ = Right$(strtext$, Len(strtext$) - InStrRev(strtext$, Chr(34) & "#000000" & Chr(34) & ">") - 10&)
    Let strtext$ = striphtml(strtext$)
    Let aimchatlinemsg$ = strtext$
End Function
Public Function aimchatlinesn() As String
    Dim strtext As String, lnglong As Long
    Let strtext$ = aimchatline$
    If strtext$ = "" Then
        Exit Function
    Else
        Let lnglong& = InStrRev(strtext$, "<br>")
        Let strtext$ = Right$(strtext$, Len(strtext$) - lnglong& - 3&)
        Let lnglong& = InStr(strtext$, "#ff0000" & Chr(34) & ">")
        Let strtext$ = Mid$(strtext$, lnglong& + Len("#ff0000" & Chr(34) & ">"), Len(strtext$) - lnglong&)
        Let strtext$ = striphtml(strtext$)
        Let aimchatlinesn$ = Left$(strtext$, InStr(strtext$, ":") - 1&)
    End If
End Function
Public Function GetText(lngwindow As Long) As String
    Dim strbuffer As String, lngtextlen As Long
    Let lngtextlen& = SendMessage(lngwindow&, WM_GETTEXTLENGTH, 0&, 0&)
    Let strbuffer$ = String(lngtextlen&, 0&)
    Call SendMessageByString(lngwindow&, WM_GETTEXT, lngtextlen& + 1&, strbuffer$)
    Let GetText$ = strbuffer$
End Function
Public Sub SetText(Window As Long, text As String)
    Call SendMessageByString(Window&, WM_SETTEXT, 0&, text$)
End Sub
Public Function striphtml(thestring As String, Optional returns As Boolean) As String
    'from juggalo32.bas

    Dim roomtext As String, takeout As String, ReplaceWith As String
    Dim whereat As Long, lefttext As String, righttext As String
    Dim takeout2 As String, whereat1 As Long, whereat2 As Long
    Dim takeout3 As String, takeout4 As String
    roomtext$ = thestring$
    If returns = True Then
        takeout$ = "<br>"
        takeout2$ = "<Br>"
        takeout3$ = "<bR>"
        takeout4$ = "<BR>"
        ReplaceWith$ = Chr(13) & Chr(10)
        whereat& = 0&
        Do: DoEvents
            whereat& = InStr(whereat& + 1, roomtext$, takeout$)
            If whereat& = 0& Then
                whereat& = InStr(whereat& + 1, roomtext$, takeout2$)
                If whereat& = 0& Then
                    whereat& = InStr(whereat& + 1, roomtext$, takeout3$)
                    If whereat& = 0& Then
                        whereat& = InStr(whereat& + 1, roomtext$, takeout4$)
                        If whereat& = 0& Then
                            Exit Do
                        End If
                    End If
                End If
            End If
            lefttext$ = Left(roomtext$, whereat& - 1)
            righttext$ = Mid(roomtext$, whereat& + 4, Len(roomtext$))
            roomtext$ = lefttext$ & ReplaceWith$ & righttext$
        Loop
    End If
    takeout$ = "<"
    takeout2$ = ">"
    whereat& = 0&
    whereat1& = 0&
    whereat2& = 0&
    Do: DoEvents
        whereat1& = InStr(whereat1& + 1, roomtext$, takeout$)
        If whereat1& = 0& Then Exit Do
        whereat2& = InStr(whereat2& + 1, roomtext$, takeout2$)
        whereat& = whereat2& - whereat1&
        lefttext$ = Left(roomtext$, whereat1& - 1)
        righttext$ = Mid(roomtext$, whereat2& + 1, Len(roomtext$) - whereat& + 1)
        roomtext$ = lefttext$ & righttext$
        whereat& = 0&
        whereat1& = 0&
        whereat2& = 0&
    Loop
   
    striphtml$ = Left(roomtext$, Len(roomtext$) - 2)

End Function

Function SendChat(message As String)
Dim aimchatwnd&
Dim wndateclass&
Dim ateclass&
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
wndateclass& = FindWindowEx(aimchatwnd&, 0&, "wndate32class", vbNullString)
wndateclass& = FindWindowEx(aimchatwnd&, wndateclass&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)

SendChat = ateclass&
Call SendMessageByString(SendChat, WM_SETTEXT, WM_CHAR, message$)

Dim oscariconbtn&
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)

SendChat = oscariconbtn&
Call leftclick(SendChat)
End Function


Public Sub newSN(sn As String, pw As String, email As String, frmname As Form)
frmname.win1.OpenURL "http://aim.aol.com/aimnew/create_new.adp?name=" & sn$ & "&password=" & pw$ & "&confirm=" & pw & "&email=" & email$ & "&promo=106712&privacy=1&pageset=Aim&client=no", 1
MsgBox ("Ok! The screen name is most likely made, if it doesnt work... the screen name was already takin, or you did something wrong... but most likely already taking")
End Sub
Function dbl()
Call sndPlaySound(App.Path + "\" & "dbl.wav", 0&)
End Function
Function sng()
Call sndPlaySound(App.Path + "\" & "sng.wav", 0&)
End Function
Function newawaymessage(message As String, newlabel As String)
Dim X&
Dim Button&
Dim oscarbuddylistwin&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)


Dim listbox&
X& = FindWindow("#32770", vbNullString)
listbox& = FindWindowEx(X&, 0&, "listbox", vbNullString)

newawaymessage = listbox&
Call runanymenubystring(oscarbuddylistwin&, "New Message...")


X& = FindWindow("#32770", vbNullString)
X& = FindWindowEx(X&, 0&, "#32770", vbNullString)
X& = FindWindowEx(X&, X&, "#32770", vbNullString)
Button& = FindWindowEx(X&, 0&, "button", vbNullString)

newawaymessage = Button&

Call leftclick(newawaymessage)
Dim ComboBox&
Dim edit&
X& = FindWindow("#32770", vbNullString)
ComboBox& = FindWindowEx(X&, 0&, "combobox", vbNullString)
edit& = FindWindowEx(ComboBox&, 0&, "edit", vbNullString)

newawaymessage = edit&
Call SendMessageByString(newawaymessage, WM_SETTEXT, WM_CHAR, newlabel$)


Dim wndateclass&
Dim ateclass&
X& = FindWindow("#32770", vbNullString)
wndateclass& = FindWindowEx(X&, 0&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)

newawaymessage = ateclass&
Call SendMessageByString(newawaymessage, WM_SETTEXT, WM_CHAR, message$)

X& = FindWindow("#32770", vbNullString)
Button& = FindWindowEx(X&, 0&, "button", vbNullString)

newawaymessage = Button&
Call leftclick(newawaymessage)

X& = FindWindow("#32770", vbNullString)
Button& = FindWindowEx(X&, 0&, "button", vbNullString)
Button& = FindWindowEx(X&, Button&, "button", vbNullString)

newawaymessage = Button&
Call leftclick(newawaymessage)
End Function
Sub INI_Write(sAppname As String, sKeyName As String, sNewString As String, sFileName As String)
Dim r As Integer
    r = WritePrivateProfileString(sAppname, sKeyName, sNewString, sFileName)
End Sub


Function INI_Read(AppName, KeyName As String, FileName As String) As String
Dim sRet As String
    sRet = String(255, Chr(0))
    INI_Read = Left(sRet, GetPrivateProfileString(AppName, ByVal KeyName, "", sRet, Len(sRet), FileName))
End Function
Public Function sendchatspiral(strstring As String)
    Dim lngindex As Long, lnglen As Long
    lnglen& = Len(strstring$)
    Do: DoEvents
        lngindex& = lngindex& + 1&
        Call SendChat(Left$(strstring$, lngindex&))
        Pause 0.6
    Loop Until lngindex& = lnglen&
End Function
Public Sub Pause(Length As Long)
    Dim Current As Long
    Let Current& = Timer
    Do Until (Timer - Current&) >= Length&
        DoEvents
    Loop
End Sub


Function openim(sn As String)

Dim oscarbuddylistwin&
Dim oscartabgroup&
Dim oscariconbtn&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
oscartabgroup& = FindWindowEx(oscarbuddylistwin&, 0&, "_oscar_tabgroup", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, 0&, "_oscar_iconbtn", vbNullString)

openim = oscariconbtn&
Call leftclick(openim)

Dim aimimessage&
Dim oscarpersistantcombo&
Dim edit&
aimimessage& = FindWindow("aim_imessage", vbNullString)
oscarpersistantcombo& = FindWindowEx(aimimessage&, 0&, "_oscar_persistantcombo", vbNullString)
edit& = FindWindowEx(oscarpersistantcombo&, 0&, "edit", vbNullString)

openim = edit&
Call SendMessageByString(openim, WM_SETTEXT, WM_CHAR, sn$)
End Function
Public Sub runanymenubystring(lngwindow As Long, strmenutext As String)
    Dim lngmmenu As Long, lngmmcount As Long, lngindex As Long
    Dim lngsubmenu As Long, lngsmcount As Long
    Dim lngindex2 As Long, lngsmid As Long, strstring As String
    Let lngmmenu& = GetMenu(lngwindow&)
    Let lngmmcount& = GetMenuItemCount(lngmmenu&)
    For lngindex& = 0& To lngmmcount& - 1&
        Let lngsubmenu& = GetSubMenu(lngmmenu&, lngindex&)
        Let lngsmcount& = GetMenuItemCount(lngsubmenu&)
        For lngindex2& = 0& To lngsmcount& - 1&
            Let lngsmid& = GetMenuItemID(lngsubmenu&, lngindex2&)
            Let strstring$ = String$(100, " ")
            Call GetMenuString(lngsubmenu&, lngsmid&, strstring$, 100&, 1&)
            If LCase$(strstring$) = ReplaceString(LCase$(strmenutext$), Chr$(0&), "") Then
                Call SendMessageLong(lngwindow&, WM_COMMAND, lngsmid&, 0&)
                Exit Sub
            End If
        Next lngindex2&
    Next lngindex&
End Sub
Function signon(sn As String, pw As String)

Dim aimcsignonwnd&
Dim ComboBox&
Dim edit&
aimcsignonwnd& = FindWindow("aim_csignonwnd", vbNullString)
ComboBox& = FindWindowEx(aimcsignonwnd&, 0&, "combobox", vbNullString)
edit& = FindWindowEx(ComboBox&, 0&, "edit", vbNullString)

signon = edit&
Call SendMessageByString(signon, WM_SETTEXT, WM_CHAR, sn$)

aimcsignonwnd& = FindWindow("aim_csignonwnd", vbNullString)
edit& = FindWindowEx(aimcsignonwnd&, 0&, "edit", vbNullString)
 
signon = edit&
Call SendMessageByString(signon, WM_SETTEXT, WM_CHAR, pw$)

Dim oscariconbtn&
aimcsignonwnd& = FindWindow("aim_csignonwnd", vbNullString)
oscariconbtn& = FindWindowEx(aimcsignonwnd&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimcsignonwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimcsignonwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)

signon = oscariconbtn&
Call leftclick(signon)
End Function
Function editprofile()
Dim X&
Dim wndateclass&
Dim ateclass&
X& = FindWindow("#32770", vbNullString)
X& = FindWindowEx(X&, 0&, "#32770", vbNullString)
wndateclass& = FindWindowEx(X&, 0&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)

editprofile = ateclass&
Call SendMessage(editprofile, SW_SHOW, 0, 0)
Call SendMessageByString(editprofile, WM_SETTEXT, WM_CHAR, "i am cool")
End Function

Function findsn() As String
    Dim lngbuddywin As Long, thecap As String
    Let lngbuddywin& = FindWindow("_oscar_buddylistwin", vbNullString)
    Let thecap$ = GetCaption(lngbuddywin&)
    If Not Right(thecap$, 20&) = "'s Buddy List Window" Then
        Let findsn$ = ""
        Exit Function
    Else
        Let findsn$ = Left$(thecap$, Len(thecap$) - 20&)
    End If
End Function
Public Function GetCaption(lngwindow As Long) As String
    Dim strbuffer As String, lngtextlen As Long
    Let lngtextlen& = GetWindowTextLength(lngwindow&)
    Let strbuffer$ = String$(lngtextlen&, 0&)
    Call GetWindowText(lngwindow&, strbuffer$, lngtextlen& + 1&)
    Let GetCaption$ = strbuffer$
End Function
Public Function ReplaceString(strstring As String, strwhat As String, strwith As String) As String
    Dim lngpos As Long
    Do While InStr(1&, strstring$, strwhat$)
        DoEvents
        Let lngpos& = InStr(1&, strstring$, strwhat$)
        Let strstring$ = Left$(strstring$, (lngpos& - 1&)) & Right$(strstring$, Len(strstring$) - (lngpos& + Len(strwhat$) - 1&))
    Loop
    Let ReplaceString$ = strstring$
End Function
Public Sub FormDrag(frmForm As Form)
    Call ReleaseCapture
    Call SendMessage(frmForm.hWnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub
Public Sub formontop(frmForm As Form, ontop As Boolean)
    If ontop = True Then Call SetWindowPos(frmForm.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
    If ontop = False Then Call SetWindowPos(frmForm.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
