﻿'Copyright Rob Latour 2016
Imports System.Reflection

Partial Public Class WindowEnterExistingPassword

    <Obfuscation(Feature:="virtualization", Exclude:=False)>
    Private Sub WindowPromptForPassword_SourceInitialized(sender As Object, e As EventArgs) Handles Me.SourceInitialized

        Me.ShowInTaskbar = False
        CheckCapsAndNumLocks()
        Me.Hide()

    End Sub

    <Obfuscation(Feature:="virtualization", Exclude:=False)>
    Private Sub WindowNewPassword_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded

        Try

            Me.Title = gEnterPasswordWindowTitle

            gEnteredPassword = String.Empty

            Dim MyHandle = SafeNativeMethods.FindWindow(Nothing, Me.Title)

            MakeTopMost(MyHandle, SafeNativeMethods.FindWindow(Nothing, Me.Title))

            Me.Show()
            PasswordBox1.Focus()

            ' The following code is used to ensure the password window comes up in such a way that the user can immediately 
            ' start typing into it; this code was inspired by http://www.xtremevbtalk.com/showthread.php?t=318187

            Dim ThreadID1 As IntPtr = SafeNativeMethods.GetWindowThreadProcessId(SafeNativeMethods.GetForegroundWindow(), 0)
            Dim ThreadID2 As IntPtr = SafeNativeMethods.GetCurrentThreadId()

            ' By sharing input state, threads share their concept of the active window
            If ThreadID1 = ThreadID2 Then
                SafeNativeMethods.BringWindowToTop(MyHandle)
                SafeNativeMethods.ShowWindow(MyHandle, SafeNativeMethods.ShowWindowCommands.SW_SHOWNORMAL)
            Else
                SafeNativeMethods.AttachThreadInput(ThreadID1, ThreadID2, True)
                SafeNativeMethods.BringWindowToTop(MyHandle)
                SafeNativeMethods.ShowWindow(MyHandle, SafeNativeMethods.ShowWindowCommands.SW_SHOWNORMAL)
                SafeNativeMethods.AttachThreadInput(ThreadID1, ThreadID2, False)
            End If

            Me.Activate()

            SeCursor(CursorState.Normal)

        Catch ex As Exception
        End Try

    End Sub

    <Obfuscation(Feature:="virtualization", Exclude:=False)>
    Private Sub WindowPromptForPassword_Closed(sender As Object, e As EventArgs) Handles Me.Closed

    End Sub

    <Obfuscation(Feature:="virtualization", Exclude:=False)>
    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnOK.Click

        On Error Resume Next

        If PasswordBox1.Password.ToString.Trim.Length > 0 Then
            gEnteredPassword = EncryptionClass.Encrypt(PasswordBox1.Password.ToString.Trim())
            Me.Close()
        Else
            PasswordBox1.Clear()
            PasswordBox1.Focus()
        End If

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub Any_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles PasswordBox1.KeyUp, Me.KeyUp

        CheckCapsAndNumLocks()

    End Sub

    Private Sub CheckCapsAndNumLocks()

        'caps lock = 20
        'num lock = 144

        If SafeNativeMethods.GetKeyState(20) Then
            Me.lblCapsLockOn.Visibility = System.Windows.Visibility.Visible
        Else
            Me.lblCapsLockOn.Visibility = System.Windows.Visibility.Hidden
        End If

        If SafeNativeMethods.GetKeyState(144) Then
            Me.lblNumLockOff.Visibility = System.Windows.Visibility.Hidden
        Else
            Me.lblNumLockOff.Visibility = System.Windows.Visibility.Visible
        End If

    End Sub

    Private Sub ThisWindow_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged

        KeepHelpOnTop()

    End Sub

End Class