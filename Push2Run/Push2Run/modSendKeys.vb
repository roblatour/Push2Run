Imports System.Drawing.Imaging
Imports System.Reflection
Imports System.Text
Imports System.Threading
Imports System.Windows.Controls.Primitives
Imports System.Windows.Forms

Module KeyboardSend

    Private Const KeyPressDelay As Integer = 0

    Private Const KEYEVENTF_EXTENDEDKEY As Integer = 1
    Private Const KEYEVENTF_KEYUP As Integer = 2

    Private Function KeyPress(ByVal vKey As VirtualKeyCode) As Boolean

        Dim ARealKeyWasSent As Boolean = False

        Static TypingDelayRequired As Boolean = False
        Static OneTimeWaitRequired As Boolean = False

        Static BuildDelayValue As Boolean = False
        Static BuildDelayValueCount As Integer = 0
        Static ValueOfBuidDelayBeingBuilt As Integer = 0

        Static OngoingDelay As Integer = 0

        If vKey = VirtualKeyCode.Invalid Then
            'do nothing

        ElseIf vKey = VirtualKeyCode.Reset Then

            TypingDelayRequired = False
            OneTimeWaitRequired = False

            OngoingDelay = 0

        ElseIf vKey = VirtualKeyCode.TypingDelay Then

            'having received this special code; the next three values passed in to this routine
            'will be used to build the typing delay time in milliseconds

            TypingDelayRequired = True
            OngoingDelay = 0

            BuildDelayValue = True
            BuildDelayValueCount = 0
            ValueOfBuidDelayBeingBuilt = 0

        ElseIf vKey = VirtualKeyCode.Wait Then

            'having received this special code; the next three values passed in to this routine
            'will be used to build the one time delay time in milliseconds

            OneTimeWaitRequired = True

            BuildDelayValue = True
            BuildDelayValueCount = 0
            ValueOfBuidDelayBeingBuilt = 0

        ElseIf BuildDelayValue Then

            'the total delay time will be built based on the next three calls into this subroutine, each call supplying a digit; 
            'first digit Is the 100s, second digit is the 10s, third digit is the 1s 

            ValueOfBuidDelayBeingBuilt = ValueOfBuidDelayBeingBuilt * 10 + CInt(vKey)
            BuildDelayValueCount += 1

            If BuildDelayValueCount = 3 Then

                If OneTimeWaitRequired Then

                    OneTimeWaitRequired = False
                    System.Threading.Thread.Sleep(ValueOfBuidDelayBeingBuilt)

                Else

                    OngoingDelay = ValueOfBuidDelayBeingBuilt

                End If

                BuildDelayValue = False
                BuildDelayValueCount = 0
                ValueOfBuidDelayBeingBuilt = 0

            End If

        Else

            KeyDown(vKey)
            KeyUp(vKey)

            ARealKeyWasSent = True

            If OngoingDelay > 0 Then
                System.Threading.Thread.Sleep(OngoingDelay)
            End If

        End If

        Return ARealKeyWasSent

    End Function

    Private Sub KeyDown(ByVal vKey As Keys)

        On Error Resume Next
        SafeNativeMethods.keybd_event(CByte(vKey), 0, KEYEVENTF_EXTENDEDKEY, 0)

    End Sub

    Private Sub KeyUp(ByVal vKey As Keys)

        On Error Resume Next
        SafeNativeMethods.keybd_event(CByte(vKey), 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)

    End Sub

    Friend Sub SendKeysToTargetWindow(ByVal KeysToSend As String)

        Try

            SetPriority(ProcessPriorityClass.High)

            Dim TransmissionString As String = UpdateWithDateAndTime(KeysToSend)

            SendTheTransmissionString(TransmissionString)

            SetPriority(ProcessPriorityClass.Normal)

        Catch ex As Exception
        End Try

        gSendingKeyesIsRequired = False

    End Sub

    
    Private Function UpdateWithDateAndTime(ByVal WorkingStringIn As String) As String

        Dim ReturnValue As String = WorkingStringIn

        Try

            Dim ws1 As String = String.Empty
            Dim ws2 As String = String.Empty
            Dim ws3 As String = String.Empty

            Dim Position1 As Int32 = ReturnValue.IndexOf("{DATETIME") + 1
            Dim Position2 As Int32 = ReturnValue.IndexOf("}", Position1) + 2

            Dim RightNow As Date = Now

            While (Position1 > 0) AndAlso (Position2 > 0)

                'get the string "{DATETIME formattedstring }"
                ws1 = Mid(ReturnValue, Position1, Position2 - Position1)

                'reduce the string to "formattedstring"
                ws2 = ws1.Remove(0, 9)
                ws2 = ws2.Remove(ws2.Length - 1, 1).Trim

                'replace "{DATETIME formattedstring }" with the formatted date/time
                ws3 = Format(RightNow, ws2)

                ReturnValue = ReturnValue.Replace(ws1, ws3)

                Position1 = ReturnValue.IndexOf("{DATETIME") + 1
                If Position1 > 0 Then Position2 = ReturnValue.IndexOf("}", Position1) + 2

            End While

        Catch ex As Exception
            ReturnValue = WorkingStringIn
        End Try

        Return ReturnValue

    End Function

    Private Sub SoftClearModifiers()

        'clears any virtual modifiers that may be left around 
        'likely not required, but done as a safeguard

        KeyboardSend.KeyUp(VirtualKeyCode.CAPSLOCK)
        KeyboardSend.KeyUp(VirtualKeyCode.SCROLLLOCK)
        KeyboardSend.KeyUp(VirtualKeyCode.NUMLOCK)
        KeyboardSend.KeyUp(VirtualKeyCode.LCONTROL)
        KeyboardSend.KeyUp(VirtualKeyCode.RCONTROL)
        KeyboardSend.KeyUp(VirtualKeyCode.LMENU)
        KeyboardSend.KeyUp(VirtualKeyCode.RMENU)
        KeyboardSend.KeyUp(VirtualKeyCode.LWIN)
        KeyboardSend.KeyUp(VirtualKeyCode.RWIN)
        KeyboardSend.KeyUp(VirtualKeyCode.LSHIFT)
        KeyboardSend.KeyUp(VirtualKeyCode.RSHIFT)

    End Sub

    
    Private Sub SendTheTransmissionString(ByVal TransmitString As String)

        ' Log(" start send the transmission string")
        Static IgnorUpcomingPrintScreen = False

        ' windows needs to be infocus and active for this to work

        Dim ShiftIsUp As Boolean = True
        Dim CtrlIsUp As Boolean = True
        Dim AltIsUp As Boolean = True
        Dim WinIsUp As Boolean = True

        Dim LastShiftKey As VirtualKeyCode = Nothing
        Dim LastCtrlKey As VirtualKeyCode = Nothing
        Dim LastAltKey As VirtualKeyCode = Nothing
        Dim LastWinKey As VirtualKeyCode = Nothing

        SoftClearModifiers()
        GetLockKeyStates()  ' This routine is handled differently in WPF (Push2Run) and Winforms (Push2Run reloader) - so it is not included in modSendKeys as modSendKeys is a shared module between Push2Run and Push2RunReloader

        Dim OriginalCAPSLock As Boolean = gCapsLock
        Dim OriginalNumbLock As Boolean = gNumbLock
        Dim OriginalScrollLock As Boolean = gScrollLock

        'Establish a common starting point of: caps off, numlock on, scrolllock off
        If OriginalCAPSLock Then KeyboardSend.KeyPress(VirtualKeyCode.CAPSLOCK)
        If Not OriginalNumbLock Then KeyboardSend.KeyPress(VirtualKeyCode.NUMLOCK)
        If OriginalScrollLock Then KeyboardSend.KeyPress(VirtualKeyCode.SCROLLLOCK)

        Try

            BuildMyKeyboard()
            KeyPress(VirtualKeyCode.Reset) ' resets delays

            Dim AllKeys() As VirtualKeyCode = GetAllKeys(TransmitString)

            'send the transmission string ...bbbb
            For Each Key As VirtualKeyCode In AllKeys

                If Key = VirtualKeyCode.Invalid Then GoTo DoneWithThisKey  ' should not get any invalid keys at this point, but keeping this here as a safeguard

                If Key = VirtualKeyCode.Release Then GoTo ClearModifierKeys

                If (Key = VirtualKeyCode.SHIFT) OrElse (Key = VirtualKeyCode.LSHIFT) OrElse (Key = VirtualKeyCode.RSHIFT) Then
                    LastShiftKey = Key
                    If ShiftIsUp Then
                        ShiftIsUp = False
                        'Console.WriteLine("Shift down")
                        KeyboardSend.KeyDown(Key)
                        GoTo DoneWithThisKey
                    End If
                End If

                If (Key = VirtualKeyCode.CONTROL) OrElse (Key = VirtualKeyCode.LCONTROL) OrElse (Key = VirtualKeyCode.RCONTROL) Then
                    LastCtrlKey = Key
                    If CtrlIsUp Then
                        CtrlIsUp = False
                        'Console.WriteLine("Ctrl down")
                        KeyboardSend.KeyDown(Key)
                        GoTo DoneWithThisKey
                    End If
                End If

                If (Key = VirtualKeyCode.MENU) OrElse (Key = VirtualKeyCode.LMENU) OrElse (Key = VirtualKeyCode.RMENU) Then
                    LastAltKey = Key
                    If AltIsUp Then
                        AltIsUp = False
                        'Console.WriteLine("Alt down")
                        KeyboardSend.KeyDown(Key)
                        GoTo DoneWithThisKey
                    End If
                End If

                If (Key = VirtualKeyCode.LWIN) OrElse (Key = VirtualKeyCode.RWIN) Then
                    LastWinKey = Key
                    If WinIsUp Then
                        WinIsUp = False
                        'Console.WriteLine("Win down")
                        KeyboardSend.KeyDown(Key)
                        GoTo DoneWithThisKey
                    End If
                End If

                If (Key = VirtualKeyCode.AltPrintScreen) Then

                    Dim ActiveWindow As IntPtr = Get_ActiveWindowHandle()

                    Dim ActiveWindowRect As SafeNativeMethods.RECT
                    SafeNativeMethods.GetClientRect(ActiveWindow, ActiveWindowRect)

                    Dim height = ActiveWindowRect.Bottom
                    Dim width = ActiveWindowRect.Right

                    Dim ActiveWindowRect2 As New System.Drawing.Rectangle()
                    SafeNativeMethods.GetWindowRect(ActiveWindow, ActiveWindowRect2)

                    ActiveWindowRect = ActiveWindowRect
                    ActiveWindowRect2 = ActiveWindowRect2

                    Dim x = ActiveWindowRect2.X
                    Dim y = ActiveWindowRect2.Y

                    Dim TargetSize As New System.Drawing.Size(width, height)

                    Dim Bitmap As New System.Drawing.Bitmap(width, height, PixelFormat.Format32bppArgb)

                    Using graphics As System.Drawing.Graphics = graphics.FromImage(Bitmap)
                        graphics.CopyFromScreen(x, y, 0, 0, TargetSize, System.Drawing.CopyPixelOperation.SourceCopy)
                    End Using


                    Dim WindowsBitmap = New System.Drawing.Bitmap(Bitmap)

                    ' the clipboard.setimage must run with an apartmentstate of sta for this to work on a separate thread
                    Dim t = New Thread(CType((Function()
                                                  Clipboard.SetImage(WindowsBitmap)
                                              End Function), ThreadStart))
                    t.SetApartmentState(ApartmentState.STA)
                    t.Start()
                    t.Join()

                    Bitmap.Dispose()
                    WindowsBitmap.Dispose()

                    IgnorUpcomingPrintScreen = True

                    GoTo DoneWithThisKey

                End If

                If (Key = VirtualKeyCode.PrintScreen) Then

                    If IgnorUpcomingPrintScreen Then

                        IgnorUpcomingPrintScreen = False

                    Else

                        Dim WindowsBitmap = New System.Drawing.Bitmap(CaptureScreen())

                        ' the clipboard.setimage must run with an apartmentstate of sta for this to work on a separate thread
                        Dim t = New Thread(CType((Function()
                                                      Clipboard.SetImage(WindowsBitmap)
                                                  End Function), ThreadStart))
                        t.SetApartmentState(ApartmentState.STA)
                        t.Start()
                        t.Join()

                        WindowsBitmap.Dispose()

                    End If

                End If

                '****************************************************************
                If KeyboardSend.KeyPress(Key) Then
                Else
                    GoTo DoneWithThisKey
                End If
                '****************************************************************

ClearModifierKeys:

                If Not CtrlIsUp Then
                    CtrlIsUp = True
                    'Console.WriteLine("Ctrl up")
                    KeyboardSend.KeyUp(LastCtrlKey)
                End If

                If Not AltIsUp Then
                    AltIsUp = True
                    'Console.WriteLine("Alt up")
                    KeyboardSend.KeyUp(LastAltKey)
                End If

                If Not WinIsUp Then
                    WinIsUp = True
                    'Console.WriteLine("Win up")
                    KeyboardSend.KeyUp(LastWinKey)
                    ' System.Threading.Thread.Sleep(1000) ' allows time for windows to react to win key combination being pressed
                End If

                If Not ShiftIsUp Then
                    ShiftIsUp = True
                    'Console.WriteLine("Shift up")
                    KeyboardSend.KeyUp(LastShiftKey)
                End If

DoneWithThisKey:

            Next

        Catch ex As Exception
        End Try

        SoftClearModifiers()

        'Restore the original Lock Key states (i.e. the Lock Key states which were in place prior to the running of this subroutine)
        GetLockKeyStates()

        If OriginalCAPSLock Then
            If gCapsLock Then
            Else
                KeyboardSend.KeyPress(VirtualKeyCode.CAPSLOCK)
            End If
        Else
            If gCapsLock Then
                KeyboardSend.KeyPress(VirtualKeyCode.CAPSLOCK)
            End If
        End If

        If gNumbLock Then
            If (Control.IsKeyLocked(Keys.NumLock)) Then
            Else
                KeyboardSend.KeyPress(VirtualKeyCode.NUMLOCK)
            End If
        Else
            If (Control.IsKeyLocked(Keys.NumLock)) Then
                KeyboardSend.KeyPress(VirtualKeyCode.NUMLOCK)
            End If
        End If

        If gScrollLock Then
            If (Control.IsKeyLocked(Keys.Scroll)) Then
            Else
                KeyboardSend.KeyPress(VirtualKeyCode.SCROLLLOCK)
            End If
        Else
            If (Control.IsKeyLocked(Keys.Scroll)) Then
                KeyboardSend.KeyPress(VirtualKeyCode.SCROLLLOCK)
            End If
        End If

        '  Log(" end send the transmission string")

    End Sub

    Private Function GetActiveWindowTitle() As String
        Const nChars As Integer = 256
        Dim handle As IntPtr = IntPtr.Zero
        Dim Buff As StringBuilder = New StringBuilder(nChars)
        handle = SafeNativeMethods.GetForegroundWindow()

        If SafeNativeMethods.GetWindowText(handle, Buff, nChars) > 0 Then
            Return Buff.ToString()
        End If

        Return Nothing
    End Function

    Friend Function Get_ActiveWindowHandle() As IntPtr

        Dim result As IntPtr

        Dim AllProcess As Process() = Process.GetProcesses()
        Dim title As String = GetActiveWindowTitle()
        Dim handle As IntPtr = SafeNativeMethods.GetForegroundWindow()
        Dim pro As Process = Nothing

        For Each p As Process In AllProcess

            Try

                If p.MainWindowTitle.Equals(title) Then
                    pro = p
                    Exit For
                End If

            Catch __unusedException1__ As Exception
            End Try
        Next

        '        Dim pi As ProcessInfo = New ProcessInfo()

        If pro Is Nothing Then
            '            pi.ProcessTitle = title
            ' pi.ProcessHandle = handle.ToInt32()
            result = handle.ToInt32()
            '           pi.ProcessName = title
        Else
            '          pi.ProcessTitle = pro.MainWindowTitle
            '         pi.ProcessName = pro.ProcessName
            ' pi.ProcessHandle = CInt(pro.MainWindowHandle)
            result = pro.MainWindowHandle

        End If

        'If title.Equals(System.Diagnostics.Process.GetCurrentProcess().MainWindowTitle) Then
        '    Dim myProcess As Process = Process.GetProcesses().Single(Function(p) p.Id <> 0 AndAlso p.Handle = handle)
        '    pi.ProcessTitle = myProcess.MainWindowTitle & "Empty"
        '    pi.ProcessHandle = handle.ToInt32()
        '    pi.ProcessName = myProcess.ProcessName & "Empty"
        'End If

        Return result

    End Function











    Friend Function CapitalizeControlKeys(ByVal InputString As String) As String

        'capitalizes all control Keys, {esc} -> {ESC}

        Dim WorkingString As String = String.Empty

        Try

            Dim CappingOn As Boolean = False

            For Each x As Char In InputString

                If x.ToString = "{" Then
                    CappingOn = True

                ElseIf x.ToString = "}" Then
                    CappingOn = False

                End If

                If CappingOn Then
                    WorkingString &= x.ToString.ToUpper
                Else
                    WorkingString &= x.ToString
                End If

            Next

        Catch ex As Exception
        End Try

        Return WorkingString

    End Function

    Friend Function GetAllKeys(ByVal InputString As String) As VirtualKeyCode()

        ' the largest the returned virtual key code array can be is 4 times the length of the input string 
        ' (1 entry for each key, 1 for an associated shift, 1 for an associated ctrl, and 1 for an associated alt)

        Dim ReturnValue(InputString.Length * 4 + 1) As VirtualKeyCode

        Dim WorkingString As String = InputString
        Dim x As Integer = 0

        While WorkingString.Length > 0

            If WorkingString.StartsWith("{}}") Then

                UpdateReturnValueWithASimpleKey("}", ReturnValue, x)
                WorkingString = WorkingString.Remove(0, 3)

            ElseIf WorkingString.StartsWith("{{}") Then

                UpdateReturnValueWithASimpleKey("{", ReturnValue, x)
                WorkingString = WorkingString.Remove(0, 3)

            ElseIf WorkingString.StartsWith("{") Then

                Dim EndMark As Integer = WorkingString.IndexOf("}")

                If EndMark = -1 Then
                    ' formatting error - ignore the { sign as it has no matching } sign
                    WorkingString = WorkingString.Remove(0, 1)
                    x -= 1

                Else

                    Dim DelayTime As Integer = 0
                    GetNamedKey(WorkingString, ReturnValue(x), DelayTime)

                    If ReturnValue(x) = VirtualKeyCode.Invalid Then

                        x -= 1

                    ElseIf (ReturnValue(x) = VirtualKeyCode.TypingDelay) OrElse (ReturnValue(x) = VirtualKeyCode.Wait) Then

                        Dim ws As String = Format(DelayTime, "000")
                        ReturnValue(x + 1) = CInt(Microsoft.VisualBasic.Mid(ws, 1, 1))
                        ReturnValue(x + 2) = CInt(Microsoft.VisualBasic.Mid(ws, 2, 1))
                        ReturnValue(x + 3) = CInt(Microsoft.VisualBasic.Mid(ws, 3, 1))
                        x += 3

                    End If

                    WorkingString = WorkingString.Remove(0, EndMark + 1)

                End If

            ElseIf WorkingString.StartsWith("}") Then

                ' formatting error - ignore the } as it was not proceeded by a matching { sign
                WorkingString = WorkingString.Remove(0, 1)
                x -= 1

            Else

                Dim SimpleKey As String = Microsoft.VisualBasic.Left(WorkingString, 1)
                UpdateReturnValueWithASimpleKey(SimpleKey, ReturnValue, x)

                WorkingString = WorkingString.Remove(0, 1)

            End If

            x += 1

        End While

        If x > 1 Then
            ReDim Preserve ReturnValue(x - 1)
        Else
            ReDim Preserve ReturnValue(0)
        End If

        Return ReturnValue

    End Function

    Friend Sub UpdateReturnValueWithASimpleKey(ByVal SimpleKey As String, ByRef ReturnValue() As VirtualKeyCode, ByRef x As Integer)

        Dim VKC As VirtualKeyCode
        Dim ShiftRequired As Boolean
        Dim AltRequired As Boolean
        Dim CtrlRequired As Boolean

        GetSimpleKeysVirtualKeyCode(SimpleKey, VKC, ShiftRequired, AltRequired, CtrlRequired)

        If VKC = VirtualKeyCode.Invalid Then

            x -= 1 ' 1 is subtracted from the index so that the invalid key will effectively be overwritten

        Else

            If CtrlRequired Then
                ReturnValue(x) = VirtualKeyCode.LCONTROL
                x += 1
            End If

            If AltRequired Then
                ReturnValue(x) = VirtualKeyCode.LMENU
                x += 1
            End If

            If ShiftRequired Then
                ReturnValue(x) = VirtualKeyCode.LSHIFT
                x += 1
            End If

            ReturnValue(x) = VKC

        End If

    End Sub

    Friend Sub GetSimpleKeysVirtualKeyCode(ByVal InputString As String, ByRef VirtKeyCode As VirtualKeyCode, ByRef ShiftRequired As Boolean, ByRef AltRequired As Boolean, ByRef CtrlRequired As Boolean)

        For x = 1 To MyKeyBoardKeys.Count - 1
            If InputString = MyKeyBoardKeys(x).UnshiftedKey Then
                ShiftRequired = False : AltRequired = False : CtrlRequired = False : VirtKeyCode = x
                Exit Sub
            End If
        Next

        For x = 1 To MyKeyBoardKeys.Count - 1
            If InputString = MyKeyBoardKeys(x).ShiftedKey Then
                ShiftRequired = True : AltRequired = False : CtrlRequired = False : VirtKeyCode = x
                Exit Sub
            End If
        Next

        For x = 1 To MyKeyBoardKeys.Count - 1
            If InputString = MyKeyBoardKeys(x).UnshiftedAltKey Then
                ShiftRequired = False : AltRequired = True : CtrlRequired = False : VirtKeyCode = x
                Exit Sub
            End If
        Next

        For x = 1 To MyKeyBoardKeys.Count - 1
            If InputString = MyKeyBoardKeys(x).ShiftedAltKey Then
                ShiftRequired = True : AltRequired = True : CtrlRequired = False : VirtKeyCode = x
                Exit Sub
            End If
        Next

        For x = 1 To MyKeyBoardKeys.Count - 1
            If InputString = MyKeyBoardKeys(x).UnShiftedCtrlKey Then
                ShiftRequired = False : AltRequired = False : CtrlRequired = True : VirtKeyCode = x
                Exit Sub
            End If
        Next

        For x = 1 To MyKeyBoardKeys.Count - 1
            If InputString = MyKeyBoardKeys(x).ShiftedCtrlKey Then
                ShiftRequired = True : AltRequired = False : CtrlRequired = True : VirtKeyCode = x
                Exit Sub
            End If
        Next

        For x = 1 To MyKeyBoardKeys.Count - 1
            If InputString = MyKeyBoardKeys(x).UnShiftedCtrlAltKey Then
                ShiftRequired = False : AltRequired = True : CtrlRequired = True : VirtKeyCode = x
                Exit Sub
            End If
        Next

        For x = 1 To MyKeyBoardKeys.Count - 1
            If InputString = MyKeyBoardKeys(x).ShiftedCtrlAltKey Then
                ShiftRequired = True : AltRequired = True : CtrlRequired = True : VirtKeyCode = x
                Exit Sub
            End If
        Next

        VirtKeyCode = 0
        ShiftRequired = False
        AltRequired = False
        CtrlRequired = False

    End Sub

    Friend Sub GetNamedKey(ByVal InputString As String, ByRef ReturnValue As VirtualKeyCode, ByRef DelayTime As Integer)

        ReturnValue = VirtualKeyCode.Invalid
        DelayTime = 0

        Try

            If InputString.Length > 0 Then

                If InputString.StartsWith("{{}") Then

                    ReturnValue = VirtualKeyCode.OEM_4

                ElseIf InputString.StartsWith("{}}") Then

                    ReturnValue = VirtualKeyCode.OEM_6

                ElseIf InputString.StartsWith("}") Then

                    ReturnValue = VirtualKeyCode.Invalid

                ElseIf InputString.StartsWith("{ALT}{PRTSC}") OrElse InputString.StartsWith("{LEFTALT}{PRTSC}") OrElse InputString.StartsWith("{RIGHTALT}{PRTSC}") Then

                    ReturnValue = VirtualKeyCode.AltPrintScreen

                ElseIf InputString.StartsWith("{PRTSC}") Then

                    ReturnValue = VirtualKeyCode.PrintScreen

                ElseIf InputString.StartsWith("{VKC") Then

                    'format should be "{VKC###}", if ok then return integer of ### otherwise return invalid

                    Dim ws As String = InputString
                    ws = ws.Remove(0, "{VKC".Length)

                    If ws.Length < 4 Then

                        ReturnValue = VirtualKeyCode.Invalid

                    Else

                        Dim wshold As String = ws
                        ws = ws.Remove(0, 3)

                        If ws.StartsWith("}") Then

                            Try

                                Dim value As Integer = CInt(Microsoft.VisualBasic.Left(wshold, 3))

                                If (value > 0) AndAlso (value < 255) Then
                                    ReturnValue = value
                                Else
                                    ReturnValue = VirtualKeyCode.Invalid
                                End If

                            Catch ex As Exception

                                ReturnValue = VirtualKeyCode.Invalid

                            End Try

                        Else

                            ReturnValue = VirtualKeyCode.Invalid

                        End If

                    End If

                ElseIf InputString.ToUpper.StartsWith("{TYPINGDELAY") Then

                    'format should be "{TYPINGDELAY###}", if ok then return integer of ### otherwise return invalid

                    Dim ws As String = InputString
                    ws = ws.Remove(0, "{TYPINGDELAY".Length)

                    If ws.Length < 4 Then

                        ReturnValue = VirtualKeyCode.Invalid

                    Else

                        Dim wshold As String = ws
                        ws = ws.Remove(0, 3)

                        If ws.StartsWith("}") Then

                            Try

                                ReturnValue = VirtualKeyCode.TypingDelay

                                Dim value As Integer = CInt(Microsoft.VisualBasic.Left(wshold, 3))
                                If (value >= 0) Then
                                    DelayTime = value
                                End If

                            Catch ex As Exception

                                ReturnValue = VirtualKeyCode.Invalid

                            End Try

                        Else

                            ReturnValue = VirtualKeyCode.Invalid

                        End If

                    End If

                ElseIf InputString.ToUpper.StartsWith("{WAIT") Then

                    'format should be "{WAIT###}", if ok then return integer of ### otherwise return invalid

                    Dim ws As String = InputString
                    ws = ws.Remove(0, "{WAIT".Length)

                    If ws.Length < 4 Then

                        ReturnValue = VirtualKeyCode.Invalid

                    Else

                        Dim wshold As String = ws
                        ws = ws.Remove(0, 3)

                        If ws.StartsWith("}") Then

                            Try

                                ReturnValue = VirtualKeyCode.Wait

                                Dim value As Integer = CInt(Microsoft.VisualBasic.Left(wshold, 3))
                                If (value >= 0) Then
                                    DelayTime = value
                                End If

                            Catch ex As Exception

                                ReturnValue = VirtualKeyCode.Invalid

                            End Try

                        Else

                            ReturnValue = VirtualKeyCode.Invalid

                        End If

                    End If

                ElseIf InputString.StartsWith("{") Then

                    Dim EndMark As Integer = InputString.IndexOf("}")
                    If EndMark > 0 Then
                        InputString = InputString.Remove(0, 1) ' remove the {
                        ReturnValue = GetKnownKeys(Microsoft.VisualBasic.Left(InputString, InputString.IndexOf("}")).ToUpper)
                    End If

                Else

                    ReturnValue = VirtualKeyCode.Invalid

                End If

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Function GetKnownKeys(ByVal CandidateKey As String) As VirtualKeyCode

        Dim VirtualKey As Integer

        Select Case CandidateKey

            Case Is = "A" : VirtualKey = VirtualKeyCode.VK_A
            Case Is = "B" : VirtualKey = VirtualKeyCode.VK_B
            Case Is = "C" : VirtualKey = VirtualKeyCode.VK_C
            Case Is = "D" : VirtualKey = VirtualKeyCode.VK_D
            Case Is = "E" : VirtualKey = VirtualKeyCode.VK_E
            Case Is = "F" : VirtualKey = VirtualKeyCode.VK_F
            Case Is = "G" : VirtualKey = VirtualKeyCode.VK_G
            Case Is = "H" : VirtualKey = VirtualKeyCode.VK_H
            Case Is = "I" : VirtualKey = VirtualKeyCode.VK_I
            Case Is = "J" : VirtualKey = VirtualKeyCode.VK_J
            Case Is = "K" : VirtualKey = VirtualKeyCode.VK_K
            Case Is = "L" : VirtualKey = VirtualKeyCode.VK_L
            Case Is = "M" : VirtualKey = VirtualKeyCode.VK_M
            Case Is = "N" : VirtualKey = VirtualKeyCode.VK_N
            Case Is = "O" : VirtualKey = VirtualKeyCode.VK_O
            Case Is = "P" : VirtualKey = VirtualKeyCode.VK_P
            Case Is = "Q" : VirtualKey = VirtualKeyCode.VK_Q
            Case Is = "R" : VirtualKey = VirtualKeyCode.VK_R
            Case Is = "S" : VirtualKey = VirtualKeyCode.VK_S
            Case Is = "T" : VirtualKey = VirtualKeyCode.VK_T
            Case Is = "U" : VirtualKey = VirtualKeyCode.VK_U
            Case Is = "V" : VirtualKey = VirtualKeyCode.VK_V
            Case Is = "W" : VirtualKey = VirtualKeyCode.VK_W
            Case Is = "X" : VirtualKey = VirtualKeyCode.VK_X
            Case Is = "Y" : VirtualKey = VirtualKeyCode.VK_Y
            Case Is = "Z" : VirtualKey = VirtualKeyCode.VK_Z

            Case Is = "1" : VirtualKey = VirtualKeyCode.VK_1
            Case Is = "2" : VirtualKey = VirtualKeyCode.VK_2
            Case Is = "3" : VirtualKey = VirtualKeyCode.VK_3
            Case Is = "4" : VirtualKey = VirtualKeyCode.VK_4
            Case Is = "5" : VirtualKey = VirtualKeyCode.VK_5
            Case Is = "6" : VirtualKey = VirtualKeyCode.VK_6
            Case Is = "7" : VirtualKey = VirtualKeyCode.VK_7
            Case Is = "8" : VirtualKey = VirtualKeyCode.VK_8
            Case Is = "9" : VirtualKey = VirtualKeyCode.VK_9
            Case Is = "0" : VirtualKey = VirtualKeyCode.VK_0

            Case Is = "ABNTC1" : VirtualKey = VirtualKeyCode.ABNT_C1
            Case Is = "ABNTC2" : VirtualKey = VirtualKeyCode.ABNT_C2
            Case Is = "ADD" : VirtualKey = VirtualKeyCode.ADD
            Case Is = "APPS" : VirtualKey = VirtualKeyCode.APPS
            Case Is = "ATTN" : VirtualKey = VirtualKeyCode.ATTN ' 
            Case Is = "BACK" : VirtualKey = VirtualKeyCode.BACK
            Case Is = "BROWSERBACK" : VirtualKey = VirtualKeyCode.BROWSER_BACK
            Case Is = "BROWSERFAVORITES" : VirtualKey = VirtualKeyCode.BROWSER_FAVORITES
            Case Is = "BROWSERFORWARD" : VirtualKey = VirtualKeyCode.BROWSER_FORWARD
            Case Is = "BROWSERHOME" : VirtualKey = VirtualKeyCode.BROWSER_HOME
            Case Is = "BROWSERREFRESH" : VirtualKey = VirtualKeyCode.BROWSER_REFRESH
            Case Is = "BROWSERSEARCH" : VirtualKey = VirtualKeyCode.BROWSER_SEARCH
            Case Is = "BROWSERSTOP" : VirtualKey = VirtualKeyCode.BROWSER_STOP
            Case Is = "CANCEL" : VirtualKey = VirtualKeyCode.CANCEL
            Case Is = "CAPSLOCK", "CAPS" : VirtualKey = VirtualKeyCode.CAPSLOCK
            Case Is = "CLEAR" : VirtualKey = VirtualKeyCode.CLEAR
            Case Is = "CRSEL" : VirtualKey = VirtualKeyCode.CRSEL
            Case Is = "DEADCHARPROCESSED" : VirtualKey = 0           ' THIS IS USUSED.  IT'S JUST HERE FOR COMPLETENESS. 
            Case Is = "DECIMAL" : VirtualKey = VirtualKeyCode.[DECIMAL]
            Case Is = "DELETE" : VirtualKey = VirtualKeyCode.DELETE
            Case Is = "DIVIDE" : VirtualKey = VirtualKeyCode.DIVIDE
            Case Is = "DOWN" : VirtualKey = VirtualKeyCode.DOWN
            Case Is = "END" : VirtualKey = VirtualKeyCode.END
            Case Is = "ERASEEOF" : VirtualKey = VirtualKeyCode.EREOF
            Case Is = "ESCAPE", "ESC" : VirtualKey = VirtualKeyCode.ESCAPE
            Case Is = "EXECUTE" : VirtualKey = VirtualKeyCode.EXECUTE
            Case Is = "EXSEL" : VirtualKey = VirtualKeyCode.EXSEL
            Case Is = "F1" : VirtualKey = VirtualKeyCode.F1
            Case Is = "F10" : VirtualKey = VirtualKeyCode.F10
            Case Is = "F11" : VirtualKey = VirtualKeyCode.F11
            Case Is = "F12" : VirtualKey = VirtualKeyCode.F12
            Case Is = "F13" : VirtualKey = VirtualKeyCode.F13
            Case Is = "F14" : VirtualKey = VirtualKeyCode.F14
            Case Is = "F15" : VirtualKey = VirtualKeyCode.F15
            Case Is = "F16" : VirtualKey = VirtualKeyCode.F16
            Case Is = "F17" : VirtualKey = VirtualKeyCode.F17
            Case Is = "F18" : VirtualKey = VirtualKeyCode.F18
            Case Is = "F19" : VirtualKey = VirtualKeyCode.F19
            Case Is = "F2" : VirtualKey = VirtualKeyCode.F2
            Case Is = "F20" : VirtualKey = VirtualKeyCode.F20
            Case Is = "F21" : VirtualKey = VirtualKeyCode.F21
            Case Is = "F22" : VirtualKey = VirtualKeyCode.F22
            Case Is = "F23" : VirtualKey = VirtualKeyCode.F23
            Case Is = "F24" : VirtualKey = VirtualKeyCode.F24
            Case Is = "F3" : VirtualKey = VirtualKeyCode.F3
            Case Is = "F4" : VirtualKey = VirtualKeyCode.F4
            Case Is = "F5" : VirtualKey = VirtualKeyCode.F5
            Case Is = "F6" : VirtualKey = VirtualKeyCode.F6
            Case Is = "F7" : VirtualKey = VirtualKeyCode.F7
            Case Is = "F8" : VirtualKey = VirtualKeyCode.F8
            Case Is = "F9" : VirtualKey = VirtualKeyCode.F9
            Case Is = "FINALMODE" : VirtualKey = VirtualKeyCode.FINAL
            Case Is = "HELP" : VirtualKey = VirtualKeyCode.HELP
            Case Is = "HOME" : VirtualKey = VirtualKeyCode.HOME
            Case Is = "IMEACCEPT" : VirtualKey = VirtualKeyCode.ACCEPT
            Case Is = "IMECONVERT" : VirtualKey = VirtualKeyCode.CONVERT
            Case Is = "IMEMODECHANGE" : VirtualKey = VirtualKeyCode.MODECHANGE
            Case Is = "IMENONCONVERT" : VirtualKey = VirtualKeyCode.NONCONVERT
            Case Is = "IMEPROCESSED" : VirtualKey = VirtualKeyCode.PROCESSKEY
            Case Is = "INSERT" : VirtualKey = VirtualKeyCode.INSERT
            Case Is = "JUNJAMODE" : VirtualKey = VirtualKeyCode.JUNJA
            Case Is = "KANAMODE" : VirtualKey = VirtualKeyCode.KANA
            Case Is = "KANJIMODE" : VirtualKey = VirtualKeyCode.KANJI
            Case Is = "LAUNCHAPPLICATION1" : VirtualKey = VirtualKeyCode.LAUNCH_APP1
            Case Is = "LAUNCHAPPLICATION2" : VirtualKey = VirtualKeyCode.LAUNCH_APP2
            Case Is = "LAUNCHMAIL" : VirtualKey = VirtualKeyCode.LAUNCH_MAIL
            Case Is = "LEFT" : VirtualKey = VirtualKeyCode.LEFT
            Case Is = "LEFTALT", "ALT" : VirtualKey = VirtualKeyCode.LMENU
            Case Is = "LEFTCTRL", "CTRL" : VirtualKey = VirtualKeyCode.LCONTROL
            Case Is = "LEFTSHIFT", "SHIFT" : VirtualKey = VirtualKeyCode.LSHIFT
            Case Is = "LEFTWIN", "WIN" : VirtualKey = VirtualKeyCode.LWIN
            Case Is = "MEDIANEXTTRACK" : VirtualKey = VirtualKeyCode.MEDIA_NEXT_TRACK
            Case Is = "MEDIAPLAYPAUSE" : VirtualKey = VirtualKeyCode.MEDIA_PLAY_PAUSE
            Case Is = "MEDIAPREVIOUSTRACK" : VirtualKey = VirtualKeyCode.MEDIA_PREV_TRACK
            Case Is = "MEDIASTOP" : VirtualKey = VirtualKeyCode.MEDIA_STOP
            Case Is = "MULTIPLY" : VirtualKey = VirtualKeyCode.MULTIPLY
            Case Is = "NEXT" : VirtualKey = VirtualKeyCode.NEXT
            Case Is = "NONAME" : VirtualKey = VirtualKeyCode.NONAME
            Case Is = "NUMLOCK" : VirtualKey = VirtualKeyCode.NUMLOCK

            Case Is = "NUMPAD0" : VirtualKey = VirtualKeyCode.NUMPAD0
            Case Is = "NUMPAD1" : VirtualKey = VirtualKeyCode.NUMPAD1
            Case Is = "NUMPAD2" : VirtualKey = VirtualKeyCode.NUMPAD2
            Case Is = "NUMPAD3" : VirtualKey = VirtualKeyCode.NUMPAD3
            Case Is = "NUMPAD4" : VirtualKey = VirtualKeyCode.NUMPAD4
            Case Is = "NUMPAD5" : VirtualKey = VirtualKeyCode.NUMPAD5
            Case Is = "NUMPAD6" : VirtualKey = VirtualKeyCode.NUMPAD6
            Case Is = "NUMPAD7" : VirtualKey = VirtualKeyCode.NUMPAD7
            Case Is = "NUMPAD8" : VirtualKey = VirtualKeyCode.NUMPAD8
            Case Is = "NUMPAD9" : VirtualKey = VirtualKeyCode.NUMPAD9

            Case Is = "OEM1" : VirtualKey = VirtualKeyCode.OEM_1
            Case Is = "OEM2" : VirtualKey = VirtualKeyCode.OEM_2
            Case Is = "OEM3" : VirtualKey = VirtualKeyCode.OEM_3
            Case Is = "OEM4" : VirtualKey = VirtualKeyCode.OEM_4
            Case Is = "OEM5" : VirtualKey = VirtualKeyCode.OEM_5
            Case Is = "OEM6" : VirtualKey = VirtualKeyCode.OEM_6
            Case Is = "OEM7" : VirtualKey = VirtualKeyCode.OEM_7
            Case Is = "OEM8" : VirtualKey = VirtualKeyCode.OEM_8
            Case Is = "OEM102" : VirtualKey = VirtualKeyCode.OEM_102

            Case Is = "OEMATTN" : VirtualKey = VirtualKeyCode.OEM_ATTN
            Case Is = "OEMAUTO" : VirtualKey = VirtualKeyCode.OEM_AUTO
            Case Is = "OEMBACKSLASH" : VirtualKey = VirtualKeyCode.OEM_102
            Case Is = "OEMBACKTAB" : VirtualKey = VirtualKeyCode.OEM_BACKTAB
            Case Is = "OEMCLEAR" : VirtualKey = VirtualKeyCode.OEM_CLEAR
            Case Is = "OEMCLOSEBRACKETS" : VirtualKey = VirtualKeyCode.OEM_6
            Case Is = "OEMCOMMA" : VirtualKey = VirtualKeyCode.OEM_COMMA
            Case Is = "OEMCOPY" : VirtualKey = VirtualKeyCode.OEM_COPY
            Case Is = "OEMENLW" : VirtualKey = VirtualKeyCode.OEM_ENLW
            Case Is = "OEMFINISH" : VirtualKey = VirtualKeyCode.OEM_FINISH
            Case Is = "OEMMINUS" : VirtualKey = VirtualKeyCode.OEM_MINUS
            Case Is = "OEMOPENBRACKETS" : VirtualKey = VirtualKeyCode.OEM_4
            Case Is = "OEMPA1" : VirtualKey = VirtualKeyCode.OEM_PA1
            Case Is = "OEMPA2" : VirtualKey = VirtualKeyCode.OEM_PA2
            Case Is = "OEMPA3" : VirtualKey = VirtualKeyCode.OEM_PA3
            Case Is = "OEMPERIOD" : VirtualKey = VirtualKeyCode.OEM_PERIOD
            Case Is = "OEMPIPE" : VirtualKey = VirtualKeyCode.OEM_5
            Case Is = "OEMPLUS" : VirtualKey = VirtualKeyCode.OEM_PLUS
            Case Is = "OEMQUESTION" : VirtualKey = VirtualKeyCode.OEM_2
            Case Is = "OEMQUOTES" : VirtualKey = VirtualKeyCode.OEM_7
            Case Is = "OEMRESET" : VirtualKey = VirtualKeyCode.OEM_RESET
            Case Is = "OEMSEMICOLON" : VirtualKey = VirtualKeyCode.OEM_1
            Case Is = "OEMTILDE" : VirtualKey = VirtualKeyCode.OEM_3
            Case Is = "OEMWSCTRL" : VirtualKey = VirtualKeyCode.OEM_WSCTRL

            Case Is = "PA1" : VirtualKey = VirtualKeyCode.PA1
            Case Is = "PACKET" : VirtualKey = VirtualKeyCode.PACKET

            Case Is = "PAUSE" : VirtualKey = VirtualKeyCode.PAUSE
            Case Is = "PLAY" : VirtualKey = VirtualKeyCode.PLAY
            Case Is = "PRINT" : VirtualKey = VirtualKeyCode.PRINT
            Case Is = "PRIOR" : VirtualKey = VirtualKeyCode.PRIOR
            Case Is = "PROCESSKEY" : VirtualKey = VirtualKeyCode.PROCESSKEY

            Case Is = "RETURN", "ENTER" : VirtualKey = VirtualKeyCode.RETURN
            Case Is = "RIGHT" : VirtualKey = VirtualKeyCode.RIGHT
            Case Is = "RIGHTALT" : VirtualKey = VirtualKeyCode.RMENU
            Case Is = "RIGHTCTRL" : VirtualKey = VirtualKeyCode.RCONTROL
            Case Is = "RIGHTSHIFT" : VirtualKey = VirtualKeyCode.RSHIFT
            Case Is = "RIGHTWIN" : VirtualKey = VirtualKeyCode.RWIN
            Case Is = "SCROLLLOCK" : VirtualKey = VirtualKeyCode.SCROLLLOCK
            Case Is = "SELECT" : VirtualKey = VirtualKeyCode.SELECT
            Case Is = "SELECTMEDIA" : VirtualKey = VirtualKeyCode.LAUNCH_MEDIA_SELECT
            Case Is = "SEPARATOR" : VirtualKey = VirtualKeyCode.SEPARATOR
            Case Is = "SLEEP" : VirtualKey = VirtualKeyCode.SLEEP
            Case Is = "SNAPSHOT" : VirtualKey = VirtualKeyCode.SNAPSHOT
            Case Is = "SPACE", " " : VirtualKey = VirtualKeyCode.SPACE
            Case Is = "SUBTRACT" : VirtualKey = VirtualKeyCode.SUBTRACT
            Case Is = "TAB" : VirtualKey = VirtualKeyCode.TAB
            Case Is = "UP" : VirtualKey = VirtualKeyCode.UP
            Case Is = "VOLUMEDOWN" : VirtualKey = VirtualKeyCode.VOLUME_DOWN
            Case Is = "VOLUMEMUTE" : VirtualKey = VirtualKeyCode.VOLUME_MUTE
            Case Is = "VOLUMEUP" : VirtualKey = VirtualKeyCode.VOLUME_UP
            Case Is = "ZOOM" : VirtualKey = VirtualKeyCode.ZOOM

            Case Is = "VKC000" : VirtualKey = VirtualKeyCode.VKC000
            Case Is = "VKC001" : VirtualKey = VirtualKeyCode.VKC001
            Case Is = "VKC002" : VirtualKey = VirtualKeyCode.VKC002
            Case Is = "VKC003" : VirtualKey = VirtualKeyCode.VKC003
            Case Is = "VKC004" : VirtualKey = VirtualKeyCode.VKC004
            Case Is = "VKC005" : VirtualKey = VirtualKeyCode.VKC005
            Case Is = "VKC006" : VirtualKey = VirtualKeyCode.VKC006
            Case Is = "VKC007" : VirtualKey = VirtualKeyCode.VKC007
            Case Is = "VKC008" : VirtualKey = VirtualKeyCode.VKC008
            Case Is = "VKC009" : VirtualKey = VirtualKeyCode.VKC009
            Case Is = "VKC010" : VirtualKey = VirtualKeyCode.VKC010
            Case Is = "VKC011" : VirtualKey = VirtualKeyCode.VKC011
            Case Is = "VKC012" : VirtualKey = VirtualKeyCode.VKC012
            Case Is = "VKC013" : VirtualKey = VirtualKeyCode.VKC013
            Case Is = "VKC014" : VirtualKey = VirtualKeyCode.VKC014
            Case Is = "VKC015" : VirtualKey = VirtualKeyCode.VKC015
            Case Is = "VKC016" : VirtualKey = VirtualKeyCode.VKC016
            Case Is = "VKC017" : VirtualKey = VirtualKeyCode.VKC017
            Case Is = "VKC018" : VirtualKey = VirtualKeyCode.VKC018
            Case Is = "VKC019" : VirtualKey = VirtualKeyCode.VKC019
            Case Is = "VKC020" : VirtualKey = VirtualKeyCode.VKC020
            Case Is = "VKC021" : VirtualKey = VirtualKeyCode.VKC021
            Case Is = "VKC022" : VirtualKey = VirtualKeyCode.VKC022
            Case Is = "VKC023" : VirtualKey = VirtualKeyCode.VKC023
            Case Is = "VKC024" : VirtualKey = VirtualKeyCode.VKC024
            Case Is = "VKC025" : VirtualKey = VirtualKeyCode.VKC025
            Case Is = "VKC026" : VirtualKey = VirtualKeyCode.VKC026
            Case Is = "VKC027" : VirtualKey = VirtualKeyCode.VKC027
            Case Is = "VKC028" : VirtualKey = VirtualKeyCode.VKC028
            Case Is = "VKC029" : VirtualKey = VirtualKeyCode.VKC029
            Case Is = "VKC030" : VirtualKey = VirtualKeyCode.VKC030
            Case Is = "VKC031" : VirtualKey = VirtualKeyCode.VKC031
            Case Is = "VKC032" : VirtualKey = VirtualKeyCode.VKC032
            Case Is = "VKC033" : VirtualKey = VirtualKeyCode.VKC033
            Case Is = "VKC034" : VirtualKey = VirtualKeyCode.VKC034
            Case Is = "VKC035" : VirtualKey = VirtualKeyCode.VKC035
            Case Is = "VKC036" : VirtualKey = VirtualKeyCode.VKC036
            Case Is = "VKC037" : VirtualKey = VirtualKeyCode.VKC037
            Case Is = "VKC038" : VirtualKey = VirtualKeyCode.VKC038
            Case Is = "VKC039" : VirtualKey = VirtualKeyCode.VKC039
            Case Is = "VKC040" : VirtualKey = VirtualKeyCode.VKC040
            Case Is = "VKC041" : VirtualKey = VirtualKeyCode.VKC041
            Case Is = "VKC042" : VirtualKey = VirtualKeyCode.VKC042
            Case Is = "VKC043" : VirtualKey = VirtualKeyCode.VKC043
            Case Is = "VKC044" : VirtualKey = VirtualKeyCode.VKC044
            Case Is = "VKC045" : VirtualKey = VirtualKeyCode.VKC045
            Case Is = "VKC046" : VirtualKey = VirtualKeyCode.VKC046
            Case Is = "VKC047" : VirtualKey = VirtualKeyCode.VKC047
            Case Is = "VKC048" : VirtualKey = VirtualKeyCode.VKC048
            Case Is = "VKC049" : VirtualKey = VirtualKeyCode.VKC049
            Case Is = "VKC050" : VirtualKey = VirtualKeyCode.VKC050
            Case Is = "VKC051" : VirtualKey = VirtualKeyCode.VKC051
            Case Is = "VKC052" : VirtualKey = VirtualKeyCode.VKC052
            Case Is = "VKC053" : VirtualKey = VirtualKeyCode.VKC053
            Case Is = "VKC054" : VirtualKey = VirtualKeyCode.VKC054
            Case Is = "VKC055" : VirtualKey = VirtualKeyCode.VKC055
            Case Is = "VKC056" : VirtualKey = VirtualKeyCode.VKC056
            Case Is = "VKC057" : VirtualKey = VirtualKeyCode.VKC057
            Case Is = "VKC058" : VirtualKey = VirtualKeyCode.VKC058
            Case Is = "VKC059" : VirtualKey = VirtualKeyCode.VKC059
            Case Is = "VKC060" : VirtualKey = VirtualKeyCode.VKC060
            Case Is = "VKC061" : VirtualKey = VirtualKeyCode.VKC061
            Case Is = "VKC062" : VirtualKey = VirtualKeyCode.VKC062
            Case Is = "VKC063" : VirtualKey = VirtualKeyCode.VKC063
            Case Is = "VKC064" : VirtualKey = VirtualKeyCode.VKC064
            Case Is = "VKC065" : VirtualKey = VirtualKeyCode.VKC065
            Case Is = "VKC066" : VirtualKey = VirtualKeyCode.VKC066
            Case Is = "VKC067" : VirtualKey = VirtualKeyCode.VKC067
            Case Is = "VKC068" : VirtualKey = VirtualKeyCode.VKC068
            Case Is = "VKC069" : VirtualKey = VirtualKeyCode.VKC069
            Case Is = "VKC070" : VirtualKey = VirtualKeyCode.VKC070
            Case Is = "VKC071" : VirtualKey = VirtualKeyCode.VKC071
            Case Is = "VKC072" : VirtualKey = VirtualKeyCode.VKC072
            Case Is = "VKC073" : VirtualKey = VirtualKeyCode.VKC073
            Case Is = "VKC074" : VirtualKey = VirtualKeyCode.VKC074
            Case Is = "VKC075" : VirtualKey = VirtualKeyCode.VKC075
            Case Is = "VKC076" : VirtualKey = VirtualKeyCode.VKC076
            Case Is = "VKC077" : VirtualKey = VirtualKeyCode.VKC077
            Case Is = "VKC078" : VirtualKey = VirtualKeyCode.VKC078
            Case Is = "VKC079" : VirtualKey = VirtualKeyCode.VKC079
            Case Is = "VKC080" : VirtualKey = VirtualKeyCode.VKC080
            Case Is = "VKC081" : VirtualKey = VirtualKeyCode.VKC081
            Case Is = "VKC082" : VirtualKey = VirtualKeyCode.VKC082
            Case Is = "VKC083" : VirtualKey = VirtualKeyCode.VKC083
            Case Is = "VKC084" : VirtualKey = VirtualKeyCode.VKC084
            Case Is = "VKC085" : VirtualKey = VirtualKeyCode.VKC085
            Case Is = "VKC086" : VirtualKey = VirtualKeyCode.VKC086
            Case Is = "VKC087" : VirtualKey = VirtualKeyCode.VKC087
            Case Is = "VKC088" : VirtualKey = VirtualKeyCode.VKC088
            Case Is = "VKC089" : VirtualKey = VirtualKeyCode.VKC089
            Case Is = "VKC090" : VirtualKey = VirtualKeyCode.VKC090
            Case Is = "VKC091" : VirtualKey = VirtualKeyCode.VKC091
            Case Is = "VKC092" : VirtualKey = VirtualKeyCode.VKC092
            Case Is = "VKC093" : VirtualKey = VirtualKeyCode.VKC093
            Case Is = "VKC094" : VirtualKey = VirtualKeyCode.VKC094
            Case Is = "VKC095" : VirtualKey = VirtualKeyCode.VKC095
            Case Is = "VKC096" : VirtualKey = VirtualKeyCode.VKC096
            Case Is = "VKC097" : VirtualKey = VirtualKeyCode.VKC097
            Case Is = "VKC098" : VirtualKey = VirtualKeyCode.VKC098
            Case Is = "VKC099" : VirtualKey = VirtualKeyCode.VKC099
            Case Is = "VKC100" : VirtualKey = VirtualKeyCode.VKC100
            Case Is = "VKC101" : VirtualKey = VirtualKeyCode.VKC101
            Case Is = "VKC102" : VirtualKey = VirtualKeyCode.VKC102
            Case Is = "VKC103" : VirtualKey = VirtualKeyCode.VKC103
            Case Is = "VKC104" : VirtualKey = VirtualKeyCode.VKC104
            Case Is = "VKC105" : VirtualKey = VirtualKeyCode.VKC105
            Case Is = "VKC106" : VirtualKey = VirtualKeyCode.VKC106
            Case Is = "VKC107" : VirtualKey = VirtualKeyCode.VKC107
            Case Is = "VKC108" : VirtualKey = VirtualKeyCode.VKC108
            Case Is = "VKC109" : VirtualKey = VirtualKeyCode.VKC109
            Case Is = "VKC110" : VirtualKey = VirtualKeyCode.VKC110
            Case Is = "VKC111" : VirtualKey = VirtualKeyCode.VKC111
            Case Is = "VKC112" : VirtualKey = VirtualKeyCode.VKC112
            Case Is = "VKC113" : VirtualKey = VirtualKeyCode.VKC113
            Case Is = "VKC114" : VirtualKey = VirtualKeyCode.VKC114
            Case Is = "VKC115" : VirtualKey = VirtualKeyCode.VKC115
            Case Is = "VKC116" : VirtualKey = VirtualKeyCode.VKC116
            Case Is = "VKC117" : VirtualKey = VirtualKeyCode.VKC117
            Case Is = "VKC118" : VirtualKey = VirtualKeyCode.VKC118
            Case Is = "VKC119" : VirtualKey = VirtualKeyCode.VKC119
            Case Is = "VKC120" : VirtualKey = VirtualKeyCode.VKC120
            Case Is = "VKC121" : VirtualKey = VirtualKeyCode.VKC121
            Case Is = "VKC122" : VirtualKey = VirtualKeyCode.VKC122
            Case Is = "VKC123" : VirtualKey = VirtualKeyCode.VKC123
            Case Is = "VKC124" : VirtualKey = VirtualKeyCode.VKC124
            Case Is = "VKC125" : VirtualKey = VirtualKeyCode.VKC125
            Case Is = "VKC126" : VirtualKey = VirtualKeyCode.VKC126
            Case Is = "VKC127" : VirtualKey = VirtualKeyCode.VKC127
            Case Is = "VKC128" : VirtualKey = VirtualKeyCode.VKC128
            Case Is = "VKC129" : VirtualKey = VirtualKeyCode.VKC129
            Case Is = "VKC130" : VirtualKey = VirtualKeyCode.VKC130
            Case Is = "VKC131" : VirtualKey = VirtualKeyCode.VKC131
            Case Is = "VKC132" : VirtualKey = VirtualKeyCode.VKC132
            Case Is = "VKC133" : VirtualKey = VirtualKeyCode.VKC133
            Case Is = "VKC134" : VirtualKey = VirtualKeyCode.VKC134
            Case Is = "VKC135" : VirtualKey = VirtualKeyCode.VKC135
            Case Is = "VKC136" : VirtualKey = VirtualKeyCode.VKC136
            Case Is = "VKC137" : VirtualKey = VirtualKeyCode.VKC137
            Case Is = "VKC138" : VirtualKey = VirtualKeyCode.VKC138
            Case Is = "VKC139" : VirtualKey = VirtualKeyCode.VKC139
            Case Is = "VKC140" : VirtualKey = VirtualKeyCode.VKC140
            Case Is = "VKC141" : VirtualKey = VirtualKeyCode.VKC141
            Case Is = "VKC142" : VirtualKey = VirtualKeyCode.VKC142
            Case Is = "VKC143" : VirtualKey = VirtualKeyCode.VKC143
            Case Is = "VKC144" : VirtualKey = VirtualKeyCode.VKC144
            Case Is = "VKC145" : VirtualKey = VirtualKeyCode.VKC145
            Case Is = "VKC146" : VirtualKey = VirtualKeyCode.VKC146
            Case Is = "VKC147" : VirtualKey = VirtualKeyCode.VKC147
            Case Is = "VKC148" : VirtualKey = VirtualKeyCode.VKC148
            Case Is = "VKC149" : VirtualKey = VirtualKeyCode.VKC149
            Case Is = "VKC150" : VirtualKey = VirtualKeyCode.VKC150
            Case Is = "VKC151" : VirtualKey = VirtualKeyCode.VKC151
            Case Is = "VKC152" : VirtualKey = VirtualKeyCode.VKC152
            Case Is = "VKC153" : VirtualKey = VirtualKeyCode.VKC153
            Case Is = "VKC154" : VirtualKey = VirtualKeyCode.VKC154
            Case Is = "VKC155" : VirtualKey = VirtualKeyCode.VKC155
            Case Is = "VKC156" : VirtualKey = VirtualKeyCode.VKC156
            Case Is = "VKC157" : VirtualKey = VirtualKeyCode.VKC157
            Case Is = "VKC158" : VirtualKey = VirtualKeyCode.VKC158
            Case Is = "VKC159" : VirtualKey = VirtualKeyCode.VKC159
            Case Is = "VKC160" : VirtualKey = VirtualKeyCode.VKC160
            Case Is = "VKC161" : VirtualKey = VirtualKeyCode.VKC161
            Case Is = "VKC162" : VirtualKey = VirtualKeyCode.VKC162
            Case Is = "VKC163" : VirtualKey = VirtualKeyCode.VKC163
            Case Is = "VKC164" : VirtualKey = VirtualKeyCode.VKC164
            Case Is = "VKC165" : VirtualKey = VirtualKeyCode.VKC165
            Case Is = "VKC166" : VirtualKey = VirtualKeyCode.VKC166
            Case Is = "VKC167" : VirtualKey = VirtualKeyCode.VKC167
            Case Is = "VKC168" : VirtualKey = VirtualKeyCode.VKC168
            Case Is = "VKC169" : VirtualKey = VirtualKeyCode.VKC169
            Case Is = "VKC170" : VirtualKey = VirtualKeyCode.VKC170
            Case Is = "VKC171" : VirtualKey = VirtualKeyCode.VKC171
            Case Is = "VKC172" : VirtualKey = VirtualKeyCode.VKC172
            Case Is = "VKC173" : VirtualKey = VirtualKeyCode.VKC173
            Case Is = "VKC174" : VirtualKey = VirtualKeyCode.VKC174
            Case Is = "VKC175" : VirtualKey = VirtualKeyCode.VKC175
            Case Is = "VKC176" : VirtualKey = VirtualKeyCode.VKC176
            Case Is = "VKC177" : VirtualKey = VirtualKeyCode.VKC177
            Case Is = "VKC178" : VirtualKey = VirtualKeyCode.VKC178
            Case Is = "VKC179" : VirtualKey = VirtualKeyCode.VKC179
            Case Is = "VKC180" : VirtualKey = VirtualKeyCode.VKC180
            Case Is = "VKC181" : VirtualKey = VirtualKeyCode.VKC181
            Case Is = "VKC182" : VirtualKey = VirtualKeyCode.VKC182
            Case Is = "VKC183" : VirtualKey = VirtualKeyCode.VKC183
            Case Is = "VKC184" : VirtualKey = VirtualKeyCode.VKC184
            Case Is = "VKC185" : VirtualKey = VirtualKeyCode.VKC185
            Case Is = "VKC186" : VirtualKey = VirtualKeyCode.VKC186
            Case Is = "VKC187" : VirtualKey = VirtualKeyCode.VKC187
            Case Is = "VKC188" : VirtualKey = VirtualKeyCode.VKC188
            Case Is = "VKC189" : VirtualKey = VirtualKeyCode.VKC189
            Case Is = "VKC190" : VirtualKey = VirtualKeyCode.VKC190
            Case Is = "VKC191" : VirtualKey = VirtualKeyCode.VKC191
            Case Is = "VKC192" : VirtualKey = VirtualKeyCode.VKC192
            Case Is = "VKC193" : VirtualKey = VirtualKeyCode.VKC193
            Case Is = "VKC194" : VirtualKey = VirtualKeyCode.VKC194
            Case Is = "VKC195" : VirtualKey = VirtualKeyCode.VKC195
            Case Is = "VKC196" : VirtualKey = VirtualKeyCode.VKC196
            Case Is = "VKC197" : VirtualKey = VirtualKeyCode.VKC197
            Case Is = "VKC198" : VirtualKey = VirtualKeyCode.VKC198
            Case Is = "VKC199" : VirtualKey = VirtualKeyCode.VKC199
            Case Is = "VKC200" : VirtualKey = VirtualKeyCode.VKC200
            Case Is = "VKC201" : VirtualKey = VirtualKeyCode.VKC201
            Case Is = "VKC202" : VirtualKey = VirtualKeyCode.VKC202
            Case Is = "VKC203" : VirtualKey = VirtualKeyCode.VKC203
            Case Is = "VKC204" : VirtualKey = VirtualKeyCode.VKC204
            Case Is = "VKC205" : VirtualKey = VirtualKeyCode.VKC205
            Case Is = "VKC206" : VirtualKey = VirtualKeyCode.VKC206
            Case Is = "VKC207" : VirtualKey = VirtualKeyCode.VKC207
            Case Is = "VKC208" : VirtualKey = VirtualKeyCode.VKC208
            Case Is = "VKC209" : VirtualKey = VirtualKeyCode.VKC209
            Case Is = "VKC210" : VirtualKey = VirtualKeyCode.VKC210
            Case Is = "VKC211" : VirtualKey = VirtualKeyCode.VKC211
            Case Is = "VKC212" : VirtualKey = VirtualKeyCode.VKC212
            Case Is = "VKC213" : VirtualKey = VirtualKeyCode.VKC213
            Case Is = "VKC214" : VirtualKey = VirtualKeyCode.VKC214
            Case Is = "VKC215" : VirtualKey = VirtualKeyCode.VKC215
            Case Is = "VKC216" : VirtualKey = VirtualKeyCode.VKC216
            Case Is = "VKC217" : VirtualKey = VirtualKeyCode.VKC217
            Case Is = "VKC218" : VirtualKey = VirtualKeyCode.VKC218
            Case Is = "VKC219" : VirtualKey = VirtualKeyCode.VKC219
            Case Is = "VKC220" : VirtualKey = VirtualKeyCode.VKC220
            Case Is = "VKC221" : VirtualKey = VirtualKeyCode.VKC221
            Case Is = "VKC222" : VirtualKey = VirtualKeyCode.VKC222
            Case Is = "VKC223" : VirtualKey = VirtualKeyCode.VKC223
            Case Is = "VKC224" : VirtualKey = VirtualKeyCode.VKC224
            Case Is = "VKC225" : VirtualKey = VirtualKeyCode.VKC225
            Case Is = "VKC226" : VirtualKey = VirtualKeyCode.VKC226
            Case Is = "VKC227" : VirtualKey = VirtualKeyCode.VKC227
            Case Is = "VKC228" : VirtualKey = VirtualKeyCode.VKC228
            Case Is = "VKC229" : VirtualKey = VirtualKeyCode.VKC229
            Case Is = "VKC230" : VirtualKey = VirtualKeyCode.VKC230
            Case Is = "VKC231" : VirtualKey = VirtualKeyCode.VKC231
            Case Is = "VKC232" : VirtualKey = VirtualKeyCode.VKC232
            Case Is = "VKC233" : VirtualKey = VirtualKeyCode.VKC233
            Case Is = "VKC234" : VirtualKey = VirtualKeyCode.VKC234
            Case Is = "VKC235" : VirtualKey = VirtualKeyCode.VKC235
            Case Is = "VKC236" : VirtualKey = VirtualKeyCode.VKC236
            Case Is = "VKC237" : VirtualKey = VirtualKeyCode.VKC237
            Case Is = "VKC238" : VirtualKey = VirtualKeyCode.VKC238
            Case Is = "VKC239" : VirtualKey = VirtualKeyCode.VKC239
            Case Is = "VKC240" : VirtualKey = VirtualKeyCode.VKC240
            Case Is = "VKC241" : VirtualKey = VirtualKeyCode.VKC241
            Case Is = "VKC242" : VirtualKey = VirtualKeyCode.VKC242
            Case Is = "VKC243" : VirtualKey = VirtualKeyCode.VKC243
            Case Is = "VKC244" : VirtualKey = VirtualKeyCode.VKC244
            Case Is = "VKC245" : VirtualKey = VirtualKeyCode.VKC245
            Case Is = "VKC246" : VirtualKey = VirtualKeyCode.VKC246
            Case Is = "VKC247" : VirtualKey = VirtualKeyCode.VKC247
            Case Is = "VKC248" : VirtualKey = VirtualKeyCode.VKC248
            Case Is = "VKC249" : VirtualKey = VirtualKeyCode.VKC249
            Case Is = "VKC250" : VirtualKey = VirtualKeyCode.VKC250
            Case Is = "VKC251" : VirtualKey = VirtualKeyCode.VKC251
            Case Is = "VKC252" : VirtualKey = VirtualKeyCode.VKC252
            Case Is = "VKC253" : VirtualKey = VirtualKeyCode.VKC253
            Case Is = "VKC254" : VirtualKey = VirtualKeyCode.VKC254

            Case Is = "RELEASE" : VirtualKey = VirtualKeyCode.Release
            Case Is = "WAIT" : VirtualKey = VirtualKeyCode.Release
            Case Is = "TYPINGDELAY" : VirtualKey = VirtualKeyCode.Release

            Case Else

                VirtualKey = VirtualKeyCode.Invalid

        End Select

        Return VirtualKey

    End Function

    'ref : http://www.kbdedit.com/manual/low_level_vk_list.html
    Friend Enum VirtualKeyCode

        LBUTTON = &H1
        RBUTTON = &H2
        CANCEL = &H3
        MBUTTON = &H4
        XBUTTON1 = &H5
        XBUTTON2 = &H6
        BACK = &H8
        TAB = &H9
        CLEAR = &HC
        [RETURN] = &HD
        SHIFT = &H10
        CONTROL = &H11
        MENU = &H12
        PAUSE = &H13
        CAPSLOCK = &H14
        KANA = &H15
        HANGEUL = &H15
        HANGUL = &H15
        JUNJA = &H17
        FINAL = &H18
        HANJA = &H19
        KANJI = &H19
        ESCAPE = &H1B
        CONVERT = &H1C
        NONCONVERT = &H1D
        ACCEPT = &H1E
        MODECHANGE = &H1F
        SPACE = &H20
        PRIOR = &H21
        [NEXT] = &H22
        [END] = &H23
        HOME = &H24
        LEFT = &H25
        UP = &H26
        RIGHT = &H27
        DOWN = &H28
        [SELECT] = &H29
        PRINT = &H2A
        EXECUTE = &H2B
        SNAPSHOT = &H2C
        INSERT = &H2D
        DELETE = &H2E
        HELP = &H2F
        VK_0 = &H30
        VK_1 = &H31
        VK_2 = &H32
        VK_3 = &H33
        VK_4 = &H34
        VK_5 = &H35
        VK_6 = &H36
        VK_7 = &H37
        VK_8 = &H38
        VK_9 = &H39
        VK_A = &H41
        VK_B = &H42
        VK_C = &H43
        VK_D = &H44
        VK_E = &H45
        VK_F = &H46
        VK_G = &H47
        VK_H = &H48
        VK_I = &H49
        VK_J = &H4A
        VK_K = &H4B
        VK_L = &H4C
        VK_M = &H4D
        VK_N = &H4E
        VK_O = &H4F
        VK_P = &H50
        VK_Q = &H51
        VK_R = &H52
        VK_S = &H53
        VK_T = &H54
        VK_U = &H55
        VK_V = &H56
        VK_W = &H57
        VK_X = &H58
        VK_Y = &H59
        VK_Z = &H5A
        LWIN = &H5B
        RWIN = &H5C
        APPS = &H5D
        SLEEP = &H5F
        NUMPAD0 = &H60
        NUMPAD1 = &H61
        NUMPAD2 = &H62
        NUMPAD3 = &H63
        NUMPAD4 = &H64
        NUMPAD5 = &H65
        NUMPAD6 = &H66
        NUMPAD7 = &H67
        NUMPAD8 = &H68
        NUMPAD9 = &H69
        MULTIPLY = &H6A
        ADD = &H6B
        SEPARATOR = &H6C
        SUBTRACT = &H6D
        [DECIMAL] = &H6E
        DIVIDE = &H6F
        F1 = &H70
        F2 = &H71
        F3 = &H72
        F4 = &H73
        F5 = &H74
        F6 = &H75
        F7 = &H76
        F8 = &H77
        F9 = &H78
        F10 = &H79
        F11 = &H7A
        F12 = &H7B
        F13 = &H7C
        F14 = &H7D
        F15 = &H7E
        F16 = &H7F
        F17 = &H80
        F18 = &H81
        F19 = &H82
        F20 = &H83
        F21 = &H84
        F22 = &H85
        F23 = &H86
        F24 = &H87
        NUMLOCK = &H90
        SCROLLLOCK = &H91
        LSHIFT = &HA0
        RSHIFT = &HA1
        LCONTROL = &HA2
        RCONTROL = &HA3
        LMENU = &HA4
        RMENU = &HA5
        BROWSER_BACK = &HA6
        BROWSER_FORWARD = &HA7
        BROWSER_REFRESH = &HA8
        BROWSER_STOP = &HA9
        BROWSER_SEARCH = &HAA
        BROWSER_FAVORITES = &HAB
        BROWSER_HOME = &HAC
        VOLUME_MUTE = &HAD
        VOLUME_DOWN = &HAE
        VOLUME_UP = &HAF
        MEDIA_NEXT_TRACK = &HB0
        MEDIA_PREV_TRACK = &HB1
        MEDIA_STOP = &HB2
        MEDIA_PLAY_PAUSE = &HB3
        LAUNCH_MAIL = &HB4
        LAUNCH_MEDIA_SELECT = &HB5
        LAUNCH_APP1 = &HB6
        LAUNCH_APP2 = &HB7
        OEM_1 = &HBA
        OEM_PLUS = &HBB
        OEM_COMMA = &HBC
        OEM_MINUS = &HBD
        OEM_PERIOD = &HBE
        OEM_2 = &HBF
        OEM_3 = &HC0
        OEM_4 = &HDB
        OEM_5 = &HDC
        OEM_6 = &HDD
        OEM_7 = &HDE
        OEM_8 = &HDF
        OEM_102 = &HE2
        PROCESSKEY = &HE5
        PACKET = &HE7
        OEM_RESET = &HE9
        OEM_PA1 = &HEB
        OEM_PA2 = &HEC
        OEM_PA3 = &HED
        OEM_WSCTRL = &HEE
        OEM_FINISH = &HF1
        OEM_COPY = &HF2
        OEM_ENLW = &HF4
        OEM_BACKTAB = &HF5
        ATTN = &HF6
        CRSEL = &HF7
        EXSEL = &HF8
        EREOF = &HF9
        PLAY = &HFA
        ZOOM = &HFB
        NONAME = &HFC
        PA1 = &HFD
        OEM_CLEAR = &HFE
        ABNT_C1 = &HC1
        ABNT_C2 = &HC1
        OEM_ATTN = &HF0
        OEM_AUTO = &HF0

        VKC000 = 0
        VKC001 = 1
        VKC002 = 2
        VKC003 = 3
        VKC004 = 4
        VKC005 = 5
        VKC006 = 6
        VKC007 = 7
        VKC008 = 8
        VKC009 = 9
        VKC010 = 10
        VKC011 = 11
        VKC012 = 12
        VKC013 = 13
        VKC014 = 14
        VKC015 = 15
        VKC016 = 16
        VKC017 = 17
        VKC018 = 18
        VKC019 = 19
        VKC020 = 20
        VKC021 = 21
        VKC022 = 22
        VKC023 = 23
        VKC024 = 24
        VKC025 = 25
        VKC026 = 26
        VKC027 = 27
        VKC028 = 28
        VKC029 = 29
        VKC030 = 30
        VKC031 = 31
        VKC032 = 32
        VKC033 = 33
        VKC034 = 34
        VKC035 = 35
        VKC036 = 36
        VKC037 = 37
        VKC038 = 38
        VKC039 = 39
        VKC040 = 40
        VKC041 = 41
        VKC042 = 42
        VKC043 = 43
        VKC044 = 44
        VKC045 = 45
        VKC046 = 46
        VKC047 = 47
        VKC048 = 48
        VKC049 = 49
        VKC050 = 50
        VKC051 = 51
        VKC052 = 52
        VKC053 = 53
        VKC054 = 54
        VKC055 = 55
        VKC056 = 56
        VKC057 = 57
        VKC058 = 58
        VKC059 = 59
        VKC060 = 60
        VKC061 = 61
        VKC062 = 62
        VKC063 = 63
        VKC064 = 64
        VKC065 = 65
        VKC066 = 66
        VKC067 = 67
        VKC068 = 68
        VKC069 = 69
        VKC070 = 70
        VKC071 = 71
        VKC072 = 72
        VKC073 = 73
        VKC074 = 74
        VKC075 = 75
        VKC076 = 76
        VKC077 = 77
        VKC078 = 78
        VKC079 = 79
        VKC080 = 80
        VKC081 = 81
        VKC082 = 82
        VKC083 = 83
        VKC084 = 84
        VKC085 = 85
        VKC086 = 86
        VKC087 = 87
        VKC088 = 88
        VKC089 = 89
        VKC090 = 90
        VKC091 = 91
        VKC092 = 92
        VKC093 = 93
        VKC094 = 94
        VKC095 = 95
        VKC096 = 96
        VKC097 = 97
        VKC098 = 98
        VKC099 = 99
        VKC100 = 100
        VKC101 = 101
        VKC102 = 102
        VKC103 = 103
        VKC104 = 104
        VKC105 = 105
        VKC106 = 106
        VKC107 = 107
        VKC108 = 108
        VKC109 = 109
        VKC110 = 110
        VKC111 = 111
        VKC112 = 112
        VKC113 = 113
        VKC114 = 114
        VKC115 = 115
        VKC116 = 116
        VKC117 = 117
        VKC118 = 118
        VKC119 = 119
        VKC120 = 120
        VKC121 = 121
        VKC122 = 122
        VKC123 = 123
        VKC124 = 124
        VKC125 = 125
        VKC126 = 126
        VKC127 = 127
        VKC128 = 128
        VKC129 = 129
        VKC130 = 130
        VKC131 = 131
        VKC132 = 132
        VKC133 = 133
        VKC134 = 134
        VKC135 = 135
        VKC136 = 136
        VKC137 = 137
        VKC138 = 138
        VKC139 = 139
        VKC140 = 140
        VKC141 = 141
        VKC142 = 142
        VKC143 = 143
        VKC144 = 144
        VKC145 = 145
        VKC146 = 146
        VKC147 = 147
        VKC148 = 148
        VKC149 = 149
        VKC150 = 150
        VKC151 = 151
        VKC152 = 152
        VKC153 = 153
        VKC154 = 154
        VKC155 = 155
        VKC156 = 156
        VKC157 = 157
        VKC158 = 158
        VKC159 = 159
        VKC160 = 160
        VKC161 = 161
        VKC162 = 162
        VKC163 = 163
        VKC164 = 164
        VKC165 = 165
        VKC166 = 166
        VKC167 = 167
        VKC168 = 168
        VKC169 = 169
        VKC170 = 170
        VKC171 = 171
        VKC172 = 172
        VKC173 = 173
        VKC174 = 174
        VKC175 = 175
        VKC176 = 176
        VKC177 = 177
        VKC178 = 178
        VKC179 = 179
        VKC180 = 180
        VKC181 = 181
        VKC182 = 182
        VKC183 = 183
        VKC184 = 184
        VKC185 = 185
        VKC186 = 186
        VKC187 = 187
        VKC188 = 188
        VKC189 = 189
        VKC190 = 190
        VKC191 = 191
        VKC192 = 192
        VKC193 = 193
        VKC194 = 194
        VKC195 = 195
        VKC196 = 196
        VKC197 = 197
        VKC198 = 198
        VKC199 = 199
        VKC200 = 200
        VKC201 = 201
        VKC202 = 202
        VKC203 = 203
        VKC204 = 204
        VKC205 = 205
        VKC206 = 206
        VKC207 = 207
        VKC208 = 208
        VKC209 = 209
        VKC210 = 210
        VKC211 = 211
        VKC212 = 212
        VKC213 = 213
        VKC214 = 214
        VKC215 = 215
        VKC216 = 216
        VKC217 = 217
        VKC218 = 218
        VKC219 = 219
        VKC220 = 220
        VKC221 = 221
        VKC222 = 222
        VKC223 = 223
        VKC224 = 224
        VKC225 = 225
        VKC226 = 226
        VKC227 = 227
        VKC228 = 228
        VKC229 = 229
        VKC230 = 230
        VKC231 = 231
        VKC232 = 232
        VKC233 = 233
        VKC234 = 234
        VKC235 = 235
        VKC236 = 236
        VKC237 = 237
        VKC238 = 238
        VKC239 = 239
        VKC240 = 240
        VKC241 = 241
        VKC242 = 242
        VKC243 = 243
        VKC244 = 244
        VKC245 = 245
        VKC246 = 246
        VKC247 = 247
        VKC248 = 248
        VKC249 = 249
        VKC250 = 250
        VKC251 = 251
        VKC252 = 252
        VKC253 = 253
        VKC254 = 254

        Invalid = -1

        Release = -2

        Reset = -3

        TypingDelay = -4

        Wait = -5

        AltPrintScreen = -6

        PrintScreen = -7

    End Enum

    Private Sub SetPriority(ByVal Prioirty As ProcessPriorityClass)

        Static Dim myProcess As Process = Process.GetCurrentProcess
        myProcess.PriorityClass = Prioirty

    End Sub

    Structure MyKeyBoard

        Dim MyKey As Integer

        Dim UnshiftedKey As String
        Dim ShiftedKey As String

        Dim UnshiftedAltKey As String
        Dim ShiftedAltKey As String

        Dim UnShiftedCtrlKey As String
        Dim ShiftedCtrlKey As String

        Dim UnShiftedCtrlAltKey As String
        Dim ShiftedCtrlAltKey As String

    End Structure

    Dim MyKeyBoardKeys(255) As MyKeyBoard

    Friend Sub BuildMyKeyboard()

        For x = 1 To 255

            Try

                With MyKeyBoardKeys(x)

                    .MyKey = x

                    .UnshiftedKey = GetKeyCodeAsString(x, False, False, False)
                    .ShiftedKey = GetKeyCodeAsString(x, True, False, False)

                    .UnshiftedAltKey = GetKeyCodeAsString(x, False, True, False)
                    .ShiftedAltKey = GetKeyCodeAsString(x, True, True, False)

                    .UnShiftedCtrlKey = GetKeyCodeAsString(x, False, False, True)
                    .ShiftedCtrlKey = GetKeyCodeAsString(x, True, False, True)

                    .UnShiftedCtrlAltKey = GetKeyCodeAsString(x, False, True, True)
                    .ShiftedCtrlAltKey = GetKeyCodeAsString(x, True, True, True)

                End With

            Catch ex As Exception

            End Try

        Next

    End Sub

    Private Function GetKeyCodeAsString(ByVal key As VirtualKeyCode, ByVal shiftkey As Boolean, ByVal altkey As Boolean, ByVal ctrlkey As Boolean) As String

        Dim keyboardState = New Byte(255) {}
        If shiftkey Then keyboardState(CInt(Keys.ShiftKey)) = &HFF
        If altkey Then keyboardState(CInt(Keys.Menu)) = &HFF
        If ctrlkey Then keyboardState(CInt(Keys.ControlKey)) = &HFF

        Dim result As System.Text.StringBuilder = New System.Text.StringBuilder(256)

        SafeNativeMethods.ToUnicode(CUInt(key), 0, keyboardState, result, 256, 0)

        Return result.ToString()

    End Function

    'Private Function GetKeyCodeAsString(ByVal key As VirtualKeyCode, ByVal shiftkey As Boolean, ByVal altkey As Boolean, ByVal ctrlkey As Boolean) As String

    '    Dim keyboardState = New Byte(255) {}
    '    If shiftkey Then keyboardState(CInt(Keys.ShiftKey)) = &HFF
    '    If altkey Then keyboardState(CInt(Keys.Menu)) = &HFF
    '    If ctrlkey Then keyboardState(CInt(Keys.ControlKey)) = &HFF

    '    Dim keyboardStateStatus As Boolean = SafeNativeMethods.GetKeyboardState(keyboardState)

    '    If Not keyboardStateStatus Then
    '        Return ""
    '    End If

    '    Dim virtualKeyCode As UInteger = CUInt(key)

    '    Dim scanCode As UInteger = SafeNativeMethods.MapVirtualKey(virtualKeyCode, 0)

    '    Dim inputLocaleIdentifier As IntPtr = SafeNativeMethods.GetKeyboardLayout(0)

    '    Dim result As System.Text.StringBuilder = New System.Text.StringBuilder()

    '    SafeNativeMethods.ToUnicodeEx(virtualKeyCode, scanCode, keyboardState, result, CInt(5), CUInt(0), inputLocaleIdentifier)

    '    Return result.ToString()

    'End Function

End Module