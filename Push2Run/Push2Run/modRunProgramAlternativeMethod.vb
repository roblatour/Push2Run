Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text


Module modRunProgramAlternativeMethod

    Friend Function RunProgramLowerPrivilege(ByVal Program As String, ByVal StartInDirectory As String, ByVal Parameters As String, ByVal ProcessingWindowStyle As ProcessWindowStyle, ByVal DesiredKeysToSend As String) As ActionStatus

        Dim ReturnValue As ActionStatus = ActionStatus.Failed

        Try

            If (Parameters.Length = 0) AndAlso (StartInDirectory.Length = 0) AndAlso (DesiredKeysToSend.Length = 0) AndAlso (ProcessingWindowStyle = ProcessWindowStyle.Normal) Then

                ReturnValue = LowerPriviledgeStart(Program, "", "")

            Else

                Dim Seperator As String = Chr(255)
                Dim AlternativeToQuotes As String = Chr(254)

                Parameters = Parameters.Replace("""", AlternativeToQuotes)

                Dim SpecialParameters = "Push2Run proxy run" & Seperator & Program & Seperator & StartInDirectory & Seperator & Parameters & Seperator & ProcessingWindowStyle.ToString & Seperator & DesiredKeysToSend
                ReturnValue = LowerPriviledgeStart("Push2RunReloader.exe", SpecialParameters, Directory.GetCurrentDirectory())

            End If

            System.Threading.Thread.Sleep(2500) ' give time for the new process to fully establish itself

        Catch ex As Exception

        End Try

        Return ReturnValue

    End Function

    Friend Function LowerPriviledgeStart(ByVal ProgramName As String, ByVal Parameters As String, ByVal WorkingDirectory As String) As ActionStatus

        Dim ReturnValue As ActionStatus = ActionStatus.Failed

        Try

            Dim currentProcess As Process = Process.GetCurrentProcess

            'Enable SeIncreaseQuotaPrivilege in this process.  (This requires administrative privileges.)
            Dim hProcessToken As IntPtr = Nothing
            OpenProcessToken(currentProcess.Handle, TOKEN_ADJUST_PRIVILEGES, hProcessToken)
            Dim tkp As TOKEN_PRIVILEGES
            tkp.PrivilegeCount = 1
            LookupPrivilegeValue(Nothing, SE_INCREASE_QUOTA_NAME, tkp.TheLuid)
            tkp.Attributes = SE_PRIVILEGE_ENABLED

            AdjustTokenPrivileges(hProcessToken, False, tkp, 0, Nothing, Nothing)

            'Get window handle representing the desktop shell.  This might not work if there is no shell window, or when
            'using a custom shell.  Also note that we're assuming that the shell is not running elevated.
            Dim hShellWnd As IntPtr = GetShellWindow()

            'Get the ID of the desktop shell process.
            Dim dwShellPID As IntPtr
            GetWindowThreadProcessId(hShellWnd, dwShellPID)

            'Open the desktop shell process in order to get the process token.
            Dim hShellProcess As IntPtr = OpenProcess(PROCESS_QUERY_INFORMATION, False, dwShellPID)
            Dim hShellProcessToken As IntPtr = Nothing
            Dim hPrimaryToken As IntPtr = Nothing

            'Get the process token of the desktop shell.
            OpenProcessToken(hShellProcess, TOKEN_DUPLICATE, hShellProcessToken)

            'Duplicate the shell's process token to get a primary token.
            Dim dwTokenRights As Integer = TOKEN_QUERY Or TOKEN_ASSIGN_PRIMARY Or TOKEN_DUPLICATE Or TOKEN_ADJUST_DEFAULT Or TOKEN_ADJUST_SESSIONID
            DuplicateTokenEx(hShellProcessToken, dwTokenRights, Nothing, SecurityImpersonation, TokenPrimary, hPrimaryToken)

            Dim si As STARTUPINFO = Nothing

            Dim pi As PROCESS_INFORMATION = Nothing

            ' si.wShowWindow = False  'testing here does not work - its the reloader that runs the program 

            si.cb = Marshal.SizeOf(si)

            Dim ptrWorkingDirectory As IntPtr = Marshal.StringToHGlobalAuto(WorkingDirectory)

            Dim sbParameters As New StringBuilder(2048)
            sbParameters.Append(Parameters)

            Dim Result As Boolean = CreateProcessWithTokenW(hPrimaryToken, 0, ProgramName, sbParameters, 0, Nothing, ptrWorkingDirectory, si, pi)

            If Result Then

                Dim LastError As Integer = Marshal.GetLastWin32Error()

                If LastError = 0 Then
                    ReturnValue = ActionStatus.Succeeded
                Else
                    ReturnValue = ActionStatus.Failed
                End If

            Else

                ReturnValue = ActionStatus.Failed

            End If

        Catch ex As Exception

        End Try

        Return ReturnValue

    End Function

    <DllImport("advapi32.dll", SetLastError:=True)>
    Private Function OpenProcessToken(ByVal ProcessHandle As IntPtr, ByVal DesiredAccess As Integer, ByRef TokenHandle As IntPtr) As Boolean
    End Function

    <DllImport("advapi32.dll", SetLastError:=True)>
    Public Function GetTokenInformation(ByVal TokenHandle As IntPtr, ByVal TokenInformationClass As TOKEN_INFORMATION_CLASS,
    ByVal TokenInformation As IntPtr, ByVal TokenInformationLength As System.UInt32,
    ByRef ReturnLength As System.UInt32) As Boolean
    End Function

    <DllImport("User32.dll", SetLastError:=True)>
    Public Function GetShellWindow() As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Function GetWindowThreadProcessId(ByVal hwnd As IntPtr,
                          ByRef lpdwProcessId As IntPtr) As Integer
    End Function

    <DllImport("kernel32.dll")>
    Private Function OpenProcess(ByVal dwDesiredAccess As UInteger, <MarshalAs(UnmanagedType.Bool)> ByVal bInheritHandle As Boolean, ByVal dwProcessId As Integer) As IntPtr
    End Function

    <DllImport("advapi32.dll", SetLastError:=True)>
    Private Function DuplicateTokenEx(
    ByVal ExistingTokenHandle As IntPtr,
    ByVal dwDesiredAccess As UInt32,
    ByRef lpThreadAttributes As SECURITY_ATTRIBUTES,
    ByVal ImpersonationLevel As Integer,
    ByVal TokenType As Integer,
    ByRef DuplicateTokenHandle As System.IntPtr) As Boolean
    End Function

    <DllImport("advapi32.dll", SetLastError:=True)>
    Private Function LookupPrivilegeValue(lpSystemName As String,
   lpName As String, ByRef lpLuid As LUID) As Boolean
    End Function

    ' Use this signature if you want the previous state information returned
    <DllImport("advapi32.dll", SetLastError:=True)>
    Private Function AdjustTokenPrivileges(
    ByVal TokenHandle As IntPtr,
    ByVal DisableAllPrivileges As Boolean,
    ByRef NewState As TOKEN_PRIVILEGES,
    ByVal BufferLengthInBytes As Integer,
    ByRef PreviousState As TOKEN_PRIVILEGES,
    ByRef ReturnLengthInBytes As Integer
  ) As Boolean
    End Function

    <DllImport("advapi32", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Function CreateProcessWithTokenW(hToken As IntPtr,
                                                   dwLogonFlags As Integer,
                                                   lpApplicationName As String,
                                                   lpCommandLine As StringBuilder,
                                                   dwCreationFlags As Integer,
                                                   lpEnvironment As IntPtr,
                                                   lpCurrentDirectory As IntPtr,
                                                   ByRef lpStartupInfo As STARTUPINFO,
                                                   ByRef lpProcessInformation As PROCESS_INFORMATION) As Boolean
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Structure SECURITY_ATTRIBUTES
        Public nLength As Integer
        Public lpSecurityDescriptor As IntPtr
        Public bInheritHandle As Integer
    End Structure

    Public Enum TOKEN_INFORMATION_CLASS
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

    Structure TOKEN_PRIVILEGES
        Public PrivilegeCount As Integer
        Public TheLuid As LUID
        Public Attributes As Integer
    End Structure

    Structure LUID
        Public LowPart As UInt32
        Public HighPart As UInt32
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Structure STARTUPINFO
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

    Structure PROCESS_INFORMATION
        Public hProcess As IntPtr
        Public hThread As IntPtr
        Public dwProcessId As Integer
        Public dwThreadId As Integer
    End Structure

    Public Const SE_PRIVILEGE_ENABLED = &H2L
    Public Const PROCESS_QUERY_INFORMATION = &H400
    Public Const TOKEN_ASSIGN_PRIMARY = &H1
    Public Const TOKEN_DUPLICATE = &H2
    Public Const TOKEN_IMPERSONATE = &H4
    Public Const TOKEN_QUERY = &H8
    Public Const TOKEN_QUERY_SOURCE = &H10
    Public Const TOKEN_ADJUST_PRIVILEGES = &H20
    Public Const TOKEN_ADJUST_GROUPS = &H40
    Public Const TOKEN_ADJUST_DEFAULT = &H80
    Public Const TOKEN_ADJUST_SESSIONID = &H100
    Public Const SecurityImpersonation = 2
    Public Const TokenPrimary = 1
    Public Const SE_INCREASE_QUOTA_NAME = "SeIncreaseQuotaPrivilege"

    <DllImport("kernel32.dll")>
    Public Function CreateProcessWithTokenW(hToken As IntPtr, dwLogonFlags As UInteger, lpApplicationName As String, lpCommandLine As StringBuilder, dwCreationFlags As UInteger, lpEnvironment As IntPtr, lpCurrentDirectory As String, ByRef lpStartupInfo As STARTUPINFO, ByRef lpProcessInformation As PROCESS_INFORMATION) As <MarshalAs(UnmanagedType.Bool)> Boolean
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


End Module
