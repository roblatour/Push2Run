
'Public NotInheritable Class Environment

'    Private Sub New()
'    End Sub

'    Friend Shared Function GetCommandLineArgs() As String
'        Throw New NotImplementedException()
'    End Function
'End Class

Imports System.Runtime.InteropServices
Imports System.Security
Imports System.Text

<SuppressUnmanagedCodeSecurityAttribute()>
Friend NotInheritable Class SafeNativeMethods

    Private Sub New()
    End Sub

    Friend Enum ShowWindowCommands
        SW_HIDE = 0
        SW_SHOWNORMAL = 1
        SW_NORMAL = 1
        SW_SHOWMINIMIZED = 2
        SW_SHOWMAXIMIZED = 3
        SW_MAXIMIZE = 3
        SW_SHOWNOACTIVATE = 4
        SW_SHOW = 5
        SW_MINIMIZE = 6
        SW_SHOWMINNOACTIVE = 7
        SW_SHOWNA = 8
        SW_RESTORE = 9
        SW_SHOWDEFAULT = 10
        SW_FORCEMINIMIZE = 11
        SW_MAX = 11
    End Enum

    Friend Structure POINTAPI
        Public x As Integer
        Public y As Integer
    End Structure

    Friend Structure RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure

    Friend Structure WINDOWPLACEMENT
        Public Length As Integer
        Public flags As Integer
        Public showCmd As Integer
        Public ptMinPosition As POINTAPI
        Public ptMaxPosition As POINTAPI
        Public rcNormalPosition As RECT
    End Structure

    Friend Const PROCESS_QUERY_INFORMATION = &H400
    Friend Const SE_INCREASE_QUOTA_NAME = "SeIncreaseQuotaPrivilege"
    Friend Const SE_PRIVILEGE_ENABLED = &H2L
    Friend Const SecurityImpersonation = 2
    Friend Const SW_SHOWMAXIMIZED As Integer = 3
    Friend Const SW_SHOWMINIMIZED As Integer = 2
    Friend Const SW_SHOWNORMAL As Integer = 1
    Friend Const TOKEN_ADJUST_DEFAULT = &H80
    Friend Const TOKEN_ADJUST_GROUPS = &H40
    Friend Const TOKEN_ADJUST_PRIVILEGES = &H20
    Friend Const TOKEN_ADJUST_SESSIONID = &H100
    Friend Const TOKEN_ASSIGN_PRIMARY = &H1
    Friend Const TOKEN_DUPLICATE = &H2
    Friend Const TOKEN_IMPERSONATE = &H4
    Friend Const TOKEN_QUERY = &H8
    Friend Const TOKEN_QUERY_SOURCE = &H10
    Friend Const TokenPrimary = 1

    <StructLayout(LayoutKind.Sequential)>
    Structure SECURITY_ATTRIBUTES
        Public nLength As Integer
        Public lpSecurityDescriptor As IntPtr
        Public bInheritHandle As Integer
    End Structure

    Friend Enum TOKEN_INFORMATION_CLASS
        TokenUser = 1
        TokenGroups
        TokenPrivileges
        TokenOwner
        TokenPrimaryGroup
        TokenDefaultDacl
        TokenSource
        TokenType
        TokenImpersonationLevel
        TokenStatistics
        TokenRestrictedSids
        TokenSessionId
        TokenGroupsAndPrivileges
        TokenSessionReference
        TokenSandBoxInert
        TokenAuditPolicy
        TokenOrigin
        TokenElevationType
        TokenLinkedToken
        TokenElevation
        TokenHasRestrictions
        TokenAccessInformation
        TokenVirtualizationAllowed
        TokenVirtualizationEnabled
        TokenIntegrityLevel
        TokenUIAccess
        TokenMandatoryPolicy
        TokenLogonSid
        MaxTokenInfoClass
    End Enum

    Friend Structure TOKEN_PRIVILEGES
        Public PrivilegeCount As Integer
        Public TheLuid As LUID
        Public Attributes As Integer
    End Structure

    Friend Structure LUID
        Public LowPart As UInt32
        Public HighPart As UInt32
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Friend Structure STARTUPINFO
        Public cb As Integer
        Public lpReserved As String
        Public lpDesktop As String
        Public lpTitle As String
        Public dwX As Integer
        Public dwY As Integer
        Public dwXSize As Integer
        Public dwYSize As Integer
        Public dwXCountChars As Integer
        Public dwYCountChars As Integer
        Public dwFillAttribute As Integer
        Public dwFlags As Integer
        Public wShowWindow As Short
        Public cbReserved2 As Short
        Public lpReserved2 As Integer
        Public hStdInput As Integer
        Public hStdOutput As Integer
        Public hStdError As Integer
    End Structure

    Friend Structure PROCESS_INFORMATION
        Public hProcess As IntPtr
        Public hThread As IntPtr
        Public dwProcessId As Integer
        Public dwThreadId As Integer
    End Structure

    Friend Declare Function AttachThreadInput Lib "user32" Alias "AttachThreadInput" (ByVal idAttach As System.UInt32, ByVal idAttachTo As System.UInt32, ByVal fAttach As Boolean) As Boolean
    Friend Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As IntPtr) As Boolean
    'Friend Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByVal lpMutexAttributes As Integer, ByVal bInitialOwner As Integer, ByVal lpName As String) As Integer
    'Friend Declare Function EnumChildWindows Lib "user32" (ByVal WindowHandle As IntPtr, ByVal Callback As EnumWindowProcess, ByVal lParam As IntPtr) As Boolean
    Friend Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    'Friend Declare Function FindWindowEx Lib "user32" (ByVal hWnd1 As IntPtr, ByVal hWnd2 As IntPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As IntPtr
    Friend Declare Function GetActiveWindow Lib "user32" Alias "GetActiveWindow" () As IntPtr
    'Friend Declare Function GetAsyncKeyState Lib "user32" (ByVal iVirtKey As Integer) As Short
    'Friend Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As IntPtr, ByVal lpClassName As StringBuilder, ByVal nMaxCount As Integer) As Integer
    Friend Declare Function GetCurrentThreadId Lib "kernel32" Alias "GetCurrentThreadId" () As IntPtr
    Friend Declare Function GetDesktopWindow Lib "user32" () As IntPtr
    'Friend Declare Function GetForegroundWindow Lib "user32" () As IntPtr
    Friend Declare Function GetForegroundWindow Lib "user32" Alias "GetForegroundWindow" () As IntPtr

    Friend Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As IntPtr, ByVal lpString As StringBuilder, ByVal cch As Integer) As Integer
    Friend Declare Function GetTopWindow Lib "user32" Alias "GetTopWindow" (ByVal hwnd As IntPtr) As IntPtr
    Friend Declare Function GetClientRect Lib "user32" Alias "GetClientRect" (ByVal hWnd As System.IntPtr, ByRef lpRECT As RECT) As Integer



    Friend Declare Function GetKeyState Lib "user32" (ByVal iVirtKey As Integer) As Short
    'Friend Declare Function GetWindow Lib "user32" (ByVal hwnd As IntPtr, ByVal wCmd As Integer) As IntPtr
    Friend Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As IntPtr, ByRef lpwndpl As WINDOWPLACEMENT) As Integer
    'Friend Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As IntPtr, ByVal lpString As StringBuilder, ByVal cch As Integer) As Integer
    'Friend Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As IntPtr) As Integer
    'Friend Declare Function GetWindowThreadProcessId Lib "user32" Alias "GetWindowThreadProcessId" (ByVal hwnd As IntPtr, ByRef lpdwProcessId As IntPtr) As IntPtr
    'Friend Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As IntPtr) As Integer
    'Friend Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Integer
    'Friend Declare Function SendMessage Lib "user32" (ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    'Friend Declare Function SendMessage Lib "user32" (ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByRef lParam As StringBuilder) As IntPtr
    'Friend Declare Function SetActiveWindow Lib "user32.dll" (ByVal hwnd As IntPtr) As IntPtr
    Friend Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As IntPtr) As Integer
    Friend Declare Function SetProcessWorkingSetSize Lib "kernel32" (ByVal process As IntPtr, ByVal minimumWorkingSetSize As IntPtr, ByVal maximumWorkingSetSize As IntPtr) As Integer
    Friend Declare Function SetWindowPos Lib "user32" (ByVal hwnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As UInt32) As Boolean
    Friend Declare Function ShowWindow Lib "user32" (ByVal hwnd As IntPtr, ByVal nCmdShow As ShowWindowCommands) As Boolean
    Friend Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Integer, ByVal dwExtraInfo As Integer)
    'Friend Delegate Function EnumWindowProcess(ByVal Handle As IntPtr, ByVal Parameter As IntPtr) As Boolean
    'Friend Delegate Function EnumWindowsProcDelegate(ByVal hWnd As Integer, ByVal lParam As Integer) As
    '

    Friend Declare Function GetWindowRect Lib "user32" (hWnd As IntPtr, ByRef lpRect As System.Drawing.Rectangle) As Boolean

    <DllImport("advapi32.dll", SetLastError:=True)>
    Friend Shared Function OpenProcessToken(ByVal ProcessHandle As IntPtr, ByVal DesiredAccess As Integer, ByRef TokenHandle As IntPtr) As Boolean
    End Function

    <DllImport("advapi32.dll", SetLastError:=True)>
    Friend Shared Function GetTokenInformation(ByVal TokenHandle As IntPtr, ByVal TokenInformationClass As TOKEN_INFORMATION_CLASS,
    ByVal TokenInformation As IntPtr, ByVal TokenInformationLength As System.UInt32,
    ByRef ReturnLength As System.UInt32) As Boolean
    End Function

    <DllImport("User32.dll", SetLastError:=True)>
    Friend Shared Function GetShellWindow() As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Friend Shared Function GetWindowThreadProcessId(ByVal hwnd As IntPtr, ByRef lpdwProcessId As IntPtr) As Integer
    End Function

    <DllImport("kernel32.dll")>
    Friend Shared Function OpenProcess(ByVal dwDesiredAccess As UInteger, <MarshalAs(UnmanagedType.Bool)> ByVal bInheritHandle As Boolean, ByVal dwProcessId As Integer) As IntPtr
    End Function

    <DllImport("advapi32.dll", SetLastError:=True)>
    Friend Shared Function DuplicateTokenEx(
    ByVal ExistingTokenHandle As IntPtr,
    ByVal dwDesiredAccess As UInt32,
    ByRef lpThreadAttributes As SECURITY_ATTRIBUTES,
    ByVal ImpersonationLevel As Integer,
    ByVal TokenType As Integer,
    ByRef DuplicateTokenHandle As System.IntPtr) As Boolean
    End Function

    <DllImport("advapi32.dll", SetLastError:=True)>
    Friend Shared Function LookupPrivilegeValue(lpSystemName As String, lpName As String, ByRef lpLuid As LUID) As Boolean
    End Function

    ' Use this signature if you want the previous state information returned
    <DllImport("advapi32.dll", SetLastError:=True)>
    Friend Shared Function AdjustTokenPrivileges(
    ByVal TokenHandle As IntPtr,
    ByVal DisableAllPrivileges As Boolean,
    ByRef NewState As TOKEN_PRIVILEGES,
    ByVal BufferLengthInBytes As Integer,
    ByRef PreviousState As TOKEN_PRIVILEGES,
    ByRef ReturnLengthInBytes As Integer
  ) As Boolean
    End Function

    <DllImport("advapi32", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Friend Shared Function CreateProcessWithTokenW(hToken As IntPtr,
                                                   dwLogonFlags As Integer,
                                                   lpApplicationName As String,
                                                   lpCommandLine As StringBuilder,
                                                   dwCreationFlags As Integer,
                                                   lpEnvironment As IntPtr,
                                                   lpCurrentDirectory As IntPtr,
                                                   ByRef lpStartupInfo As STARTUPINFO,
                                                   ByRef lpProcessInformation As PROCESS_INFORMATION) As Boolean
    End Function


    <DllImport("kernel32.dll")>
    Public Shared Function CreateProcessWithTokenW(hToken As IntPtr, dwLogonFlags As UInteger, lpApplicationName As String, lpCommandLine As StringBuilder, dwCreationFlags As UInteger, lpEnvironment As IntPtr, lpCurrentDirectory As String, ByRef lpStartupInfo As STARTUPINFO, ByRef lpProcessInformation As PROCESS_INFORMATION) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function


    <DllImport("user32.dll")>
    Public Shared Function GetKeyboardState(ByVal lpKeyState As Byte()) As Boolean
    End Function

    <DllImport("user32.dll")>
    Public Shared Function MapVirtualKey(ByVal uCode As UInteger, ByVal uMapType As UInteger) As UInteger
    End Function

    <DllImport("user32.dll")>
    Public Shared Function GetKeyboardLayout(ByVal idThread As UInteger) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Public Shared Function ToUnicodeEx(ByVal wVirtKey As UInteger, ByVal wScanCode As UInteger, ByVal lpKeyState As Byte(),
    <Out, MarshalAs(UnmanagedType.LPWStr)> ByVal pwszBuff As StringBuilder, ByVal cchBuff As Integer, ByVal wFlags As UInteger, ByVal dwhkl As IntPtr) As Integer
    End Function

    'Shared Function CreateProcess(lpApplicationName As IntPtr,
    '                              lpCommandLine As IntPtr,
    '                              ByRef lpProcessAttributes As SECURITY_ATTRIBUTES,
    '                              ByRef lpThreadAttributes As SECURITY_ATTRIBUTES,
    '                              bInheritHandles As Boolean,
    '                              dwCreationFlags As UInt32,
    '                              lpEnvironment As IntPtr,
    '                              lpCurrentDirectory As StringBuilder,
    '                              <[In]()> ByRef lpStartupInfo As STARTUPINFO,
    '                              <[Out]()> ByRef lpProcessInformation As PROCESS_INFORMATION) As Boolean
    'End Function


    <DllImport("user32.dll")>
    Public Shared Function ToUnicode(ByVal virtualKeyCode As UInteger, ByVal scanCode As UInteger, ByVal keyboardState As Byte(),
    <Out, MarshalAs(UnmanagedType.LPWStr, SizeConst:=64)> ByVal receivingBuffer As StringBuilder, ByVal bufferSize As Integer, ByVal flags As UInteger) As Integer
    End Function


End Class


