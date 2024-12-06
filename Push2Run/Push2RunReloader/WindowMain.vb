'Copyright Rob Latour 2024

Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows.Automation

Public Class frmReloader

    ' ref https://stackoverflow.com/questions/48125039/vb-net-convert-string-to-lpwstr

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try

            myFormLibrary.WindowMain = Me
            GetLockKeyStates()

            Dim Seperator As String = Chr(255)

            Dim CommandLine As String = Microsoft.VisualBasic.Command.ToString.Trim

            If CommandLine.StartsWith("StartAdmin") OrElse (CommandLine.StartsWith("RestartAdmin")) Then

                System.Threading.Thread.Sleep(2500) ' give time for the current running program to shut itself down 

                Dim StartupInfo As New ProcessStartInfo

                With StartupInfo

                    .CreateNoWindow = True
                    .UseShellExecute = True
                    .WorkingDirectory = System.Environment.CurrentDirectory
                    .FileName = "Push2Run.exe"

                    If CommandLine.StartsWith("StartAdmin") Then
                        .Arguments = "StartAdmin" & Seperator
                    Else
                        .Arguments = "RestartAdmin" & Seperator
                    End If

                    .WindowStyle = ProcessWindowStyle.Minimized
                    .Verb = "runas"

                End With

                Process.Start(StartupInfo)

            ElseIf CommandLine.StartsWith("StartNormal") OrElse (CommandLine.StartsWith("RestartNormal")) Then

                System.Threading.Thread.Sleep(2500) ' give time for the current running program to shut itself down 

                Dim ProgramName As String = System.Environment.CurrentDirectory & "\" & "Push2Run.exe"

                Dim Arguments As String = String.Empty
                If CommandLine.StartsWith("StartNormal") Then
                    Arguments = "StartNormal" & Seperator
                Else
                    Arguments = "RestartNormal" & Seperator
                End If

                LowerPriviledgeStart(ProgramName, Arguments, System.Environment.CurrentDirectory)

            ElseIf CommandLine.StartsWith("Recover") Then

                System.Threading.Thread.Sleep(2500) ' give time for the current running program to shut itself down 'v3.6.2 changed from 2000 to 2500

                Dim ProgramName As String = System.Environment.CurrentDirectory & "\" & "Push2Run.exe"

                Dim Arguments As String = "StartNormal"

                SamePriviledgeStart(ProgramName, Arguments, System.Environment.CurrentDirectory)

            Else

                RegisterForAutomationEvents()

                Dim CommandLineArray() As String = Environment.GetCommandLineArgs

                CommandLine = String.Empty
                For Each entry In CommandLineArray
                    CommandLine &= " " & entry
                Next

                CommandLine = CommandLine.Trim

                If CommandLine.ToUpper.StartsWith("PUSH2RUN PROXY RUN") Then

                    Dim AlternativeToQuotes As String = Chr(254)

                    Dim args() = CommandLine.Split(Seperator)

                    Dim Program As String = String.Empty
                    Dim StartInDirectory As String = String.Empty
                    Dim Parameters As String = String.Empty
                    Dim RequestedWindowStyle As ProcessWindowStyle = ProcessWindowStyle.Normal
                    Dim KeysToSend As String = String.Empty

                    If args.Length > 1 Then Program = args(1).Trim

                    If args.Length > 2 Then StartInDirectory = args(2).Trim

                    If args.Length > 3 Then Parameters = args(3).Trim
                    ' sending data via LowerPriviledgeStart from the boss causes the quotes around the parms to be lost, however chr(254) is not lost.  
                    ' so the quotes are replaced by chr(254) in the boss, and then later restored here in this code so that the program can be run correctly
                    Parameters = Parameters.Replace(AlternativeToQuotes, """")

                    If args.Length > 4 Then

                        Dim ws As String = args(4).Trim.ToLower
                        Select Case ws

                            Case Is = "minimized"
                                RequestedWindowStyle = ProcessWindowStyle.Minimized

                            Case Is = "normal"
                                RequestedWindowStyle = ProcessWindowStyle.Normal

                            Case Is = "maximized"
                                RequestedWindowStyle = ProcessWindowStyle.Maximized

                            Case Is = "hidden"
                                RequestedWindowStyle = ProcessWindowStyle.Hidden

                            Case Else
                                RequestedWindowStyle = ProcessWindowStyle.Normal

                        End Select

                    End If

                    If args.Length > 5 Then KeysToSend = args(5).Trim

                    If Program.ToUpper.Trim = "DESKTOP" Then

                        RunDesktop(KeysToSend)

                    Else

                        RunProgramStandard(Program, StartInDirectory, Parameters, RequestedWindowStyle, KeysToSend)

                    End If

                End If

            End If

            System.Threading.Thread.Sleep(2500) ' give time for the new process to fully establish itself

        Catch ex As Exception

        End Try

        Me.Close()

    End Sub

    Friend Sub RunDesktop(ByVal DesiredKeysToSend As String)

        Dim Handle As IntPtr = SafeNativeMethods.GetDesktopWindow()
        SafeNativeMethods.SetForegroundWindow(Handle)

        gDesiredKeysToSend = DesiredKeysToSend

        SendKeysToTargetWindow(gDesiredKeysToSend)

        gDesiredKeysToSend = String.Empty

    End Sub

    Public Sub Safely_GetKeylockStates()
        Call Me.BeginInvoke(Me.CallSafelyGetKeyLockStates)
    End Sub
    Dim CallSafelyGetKeyLockStates As New MethodInvoker(AddressOf Me.CallSafelyGetKeyLockStates_Private)
    Private Sub CallSafelyGetKeyLockStates_Private()

        gCapsLock = My.Computer.Keyboard.CapsLock
        gNumbLock = My.Computer.Keyboard.NumLock
        gScrollLock = My.Computer.Keyboard.ScrollLock

    End Sub

    Friend Sub RunProgramStandard(ByVal Program As String, ByVal StartInDirectory As String, ByVal Parameters As String, ByVal WindowProcessingStyle As ProcessWindowStyle, ByVal DesiredKeysToSend As String)

        Dim ReturnValue As Boolean = False

        Try

            Dim myProcess As New Process

            With myProcess.StartInfo

                .FileName = Program

                If StartInDirectory.Length > 0 Then .WorkingDirectory = StartInDirectory

                If Parameters.Length > 0 Then .Arguments = Parameters

                .Verb = "open"

                .CreateNoWindow = True

                Select Case WindowProcessingStyle

                    Case Is = WindowProcessingStyle.Minimized
                        gDesiredWindowState = WindowVisualState.Minimized
                        .WindowStyle = ProcessWindowStyle.Minimized

                    Case Is = WindowProcessingStyle.Normal
                        gDesiredWindowState = WindowVisualState.Normal
                        .WindowStyle = ProcessWindowStyle.Normal

                    Case Is = WindowProcessingStyle.Maximized
                        gDesiredWindowState = WindowVisualState.Maximized
                        .WindowStyle = ProcessWindowStyle.Maximized

                    Case Else
                        .WindowStyle = ProcessWindowStyle.Hidden
                        gDesiredWindowState = Nothing

                End Select

                .UseShellExecute = True

            End With

            gProcessStartTime = Now ' allows automation to watch and react to a new window
            gDesiredKeysToSend = DesiredKeysToSend

            Dim ProcessStarted As Boolean = myProcess.Start()

            ReturnValue = True

        Catch ex As Exception

        End Try

    End Sub

    Private Sub LowerPriviledgeStart(ByVal ProgramName As String, ByVal Parmaters As String, ByVal WorkingDirectory As String)

        Try

            Dim currentProcess As Process = Process.GetCurrentProcess

            'Enable SeIncreaseQuotaPrivilege in this process.  (This requires administrative privileges.)
            Dim hProcessToken As IntPtr = Nothing
            SafeNativeMethods.OpenProcessToken(currentProcess.Handle, SafeNativeMethods.TOKEN_ADJUST_PRIVILEGES, hProcessToken)
            Dim tkp As SafeNativeMethods.TOKEN_PRIVILEGES
            tkp.PrivilegeCount = 1
            SafeNativeMethods.LookupPrivilegeValue(Nothing, SafeNativeMethods.SE_INCREASE_QUOTA_NAME, tkp.TheLuid)
            tkp.Attributes = SafeNativeMethods.SE_PRIVILEGE_ENABLED

            SafeNativeMethods.AdjustTokenPrivileges(hProcessToken, False, tkp, 0, Nothing, Nothing)

            'Get window handle representing the desktop shell.  This might not work if there is no shell window, or when
            'using a custom shell.  Also note that we're assuming that the shell is not running elevated.
            Dim hShellWnd As IntPtr = SafeNativeMethods.GetShellWindow()

            'Get the ID of the desktop shell process.
            Dim dwShellPID As IntPtr
            SafeNativeMethods.GetWindowThreadProcessId(hShellWnd, dwShellPID)

            'Open the desktop shell process in order to get the process token.
            Dim hShellProcess As IntPtr = SafeNativeMethods.OpenProcess(SafeNativeMethods.PROCESS_QUERY_INFORMATION, False, dwShellPID)
            Dim hShellProcessToken As IntPtr = Nothing
            Dim hPrimaryToken As IntPtr = Nothing

            'Get the process token of the desktop shell.
            SafeNativeMethods.OpenProcessToken(hShellProcess, SafeNativeMethods.TOKEN_DUPLICATE, hShellProcessToken)

            'Duplicate the shell's process token to get a primary token.
            Dim dwTokenRights As Integer = SafeNativeMethods.TOKEN_QUERY Or SafeNativeMethods.TOKEN_ASSIGN_PRIMARY Or SafeNativeMethods.TOKEN_DUPLICATE Or SafeNativeMethods.TOKEN_ADJUST_DEFAULT Or SafeNativeMethods.TOKEN_ADJUST_SESSIONID
            SafeNativeMethods.DuplicateTokenEx(hShellProcessToken, dwTokenRights, Nothing, SafeNativeMethods.SecurityImpersonation, SafeNativeMethods.TokenPrimary, hPrimaryToken)

            Dim si As SafeNativeMethods.STARTUPINFO = Nothing

            Dim pi As SafeNativeMethods.PROCESS_INFORMATION = Nothing

            si.cb = Marshal.SizeOf(si)

            Dim ptrWorkingDirectory As IntPtr = Marshal.StringToHGlobalAuto(WorkingDirectory)

            Dim sbParameters As New StringBuilder(2048)
            sbParameters.Append(Parmaters)

            Dim Result As Boolean = SafeNativeMethods.CreateProcessWithTokenW(hPrimaryToken, 0, ProgramName, sbParameters, 0, Nothing, ptrWorkingDirectory, si, pi)

        Catch ex As Exception

        End Try

    End Sub

    'v3.6.2 only applies to when Push2Run needs to be restarted following the recovery of a corrupt setting fils
    Private Sub SamePriviledgeStart(ByVal ProgramName As String, ByVal Parmaters As String, ByVal WorkingDirectory As String)

        Try

            Dim myProcess As New Process

            With myProcess.StartInfo

                .CreateNoWindow = True
                .UseShellExecute = True
                .WorkingDirectory = WorkingDirectory
                .FileName = ProgramName
                .Verb = "open"
                .Arguments = Parmaters

            End With

            Dim ProcessStarted As Boolean = myProcess.Start()

            System.Threading.Thread.Sleep(2000)

        Catch ex As Exception

            MsgBox("Gnats 2 :" & vbCrLf & ex.ToString)

        End Try

    End Sub


#Region "Windows Automation"

    'ref: https://docs.microsoft.com/en-us/dotnet/api/system.windows.automation.windowpattern.windowopenedevent?view=netframework-4.7.2
    'ref: https://stackoverflow.com/questions/54120120/getting-process-start-with-windowstyle-to-work.
    'ref: https://docs.microsoft.com/en-us/dotnet/api/system.windows.automation.windowpattern.setwindowvisualstate?view=netframework-4.7.2

    'setup ...
    'RegisterForAutomationEvents() 

    Friend gDesiredWindowState As WindowVisualState
    Friend gDesiredKeysToSend As String = String.Empty
    Friend gProcessStartTime As DateTime = Now.AddDays(-1)

    Private ReadOnly WindowOpenedEvent As AutomationEvent
    Friend Sub RegisterForAutomationEvents()

        Dim eventHandler As AutomationEventHandler = AddressOf OnWindowOpen
        Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent, AutomationElement.RootElement, TreeScope.Children, eventHandler)

    End Sub

    Dim LockingObject = New Object

    Private Sub OnWindowOpen(ByVal src As Object, ByVal e As AutomationEventArgs)

        ' can't just test for processor id matching the launched program as the window that opens 
        ' doesn't always have the same pid as was started in the myprocess.start 
        ' (the windows calculator is an example of this)

        ' the following code will look for a new window to be opened in 5 seconds or less of push2run launching a program
        ' if one does't get opened in that time frame it will be ignored

        SyncLock LockingObject 'test if synclock closes when sube exits

            Try


                If Now > gProcessStartTime.AddSeconds(5) Then GoTo AllDone

                gProcessStartTime.AddDays(-1)

                Dim OKToSendKeys As Boolean = False

                Dim Handle As IntPtr

                Try

                    If (gDesiredWindowState = WindowVisualState.Minimized) OrElse
                       (gDesiredWindowState = WindowVisualState.Normal) OrElse
                       (gDesiredWindowState = WindowVisualState.Maximized) Then

                        Dim SourceElement As AutomationElement = DirectCast(src, AutomationElement)
                        Handle = SourceElement.Current.NativeWindowHandle

                        Dim WindowPattern As WindowPattern = DirectCast(SourceElement.GetCurrentPattern(WindowPattern.Pattern), WindowPattern)

                        Handle = SourceElement.Current.NativeWindowHandle

                        If WindowPattern.WaitForInputIdle(10000) Then
                        Else
                            Exit Try 'object is not responding
                        End If

                        Threading.Thread.Sleep(100)

                        If WindowPattern.Current.IsModal Then

                        Else

                            If (WindowPattern.Current.CanMinimize) AndAlso (gDesiredWindowState = WindowVisualState.Minimized) Then
                                WindowPattern.SetWindowVisualState(WindowVisualState.Minimized)
                                Exit Try
                            End If

                            If gDesiredWindowState = WindowVisualState.Normal Then
                                WindowPattern.SetWindowVisualState(WindowVisualState.Normal)
                                SourceElement.SetFocus()
                                OKToSendKeys = True
                                Exit Try
                            End If

                            If (WindowPattern.Current.CanMaximize) AndAlso (gDesiredWindowState = WindowVisualState.Maximized) Then
                                WindowPattern.SetWindowVisualState(WindowVisualState.Maximized)
                                SourceElement.SetFocus()
                                OKToSendKeys = True
                                Exit Try
                            End If

                        End If

                    End If

                Catch ex As Exception
                End Try


                If OKToSendKeys AndAlso (gDesiredKeysToSend.Length > 0) Then

                    Threading.Thread.Sleep(1250)
                    SendKeysToTargetWindow(gDesiredKeysToSend)
                    'gDesiredKeysToSend = String.Empty

                End If

                'gDesiredWindowState = Nothing

            Catch ex As Exception

            End Try

AllDone:

            gDesiredKeysToSend = String.Empty
            gDesiredWindowState = Nothing

        End SyncLock


    End Sub

#End Region

End Class
