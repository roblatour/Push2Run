Imports System.Reflection

Partial Public Class WindowChangePassword

    <Obfuscation(Feature:="virtualization", Exclude:=False)>
    Private Sub WindowNewPassword_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded

        Me.ShowInTaskbar = False

        MakeTopMost(SafeNativeMethods.FindWindow(Nothing, Me.Title), My.Settings.AlwaysOnTop)

        CheckCapsAndNumLocks()
        OldPasswordBox.Focus()

        ResetEncryptionAndDecriptionToReadAndWrite(ResetEncryptionDecriptionLevel.Passwords)

        SeCursor(CursorState.Normal)

    End Sub

    <Obfuscation(Feature:="virtualization", Exclude:=False)>
    Private Sub WindowChangePassword_Closing(sender As Object, e As ComponentModel.CancelEventArgs) Handles Me.Closing

        ResetEncryptionAndDecriptionToReadAndWrite(ResetEncryptionDecriptionLevel.Data)

    End Sub

    <Obfuscation(Feature:="virtualization", Exclude:=False)>
    Private Sub Any_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles WindowNewPasswordBox1.KeyUp, WindowNewPasswordBox2.KeyUp, OldPasswordBox.KeyUp, Me.KeyUp

        CheckCapsAndNumLocks()

    End Sub

    <Obfuscation(Feature:="virtualization", Exclude:=False)>
    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnOK.Click

        Dim OldPlainTextPassword As String = OldPasswordBox.Password.ToString.Trim
        Dim NewPlainTextPassword As String = WindowNewPasswordBox1.Password.ToString.Trim
        Dim ConfirmedPlainTextPassword As String = WindowNewPasswordBox2.Password.ToString.Trim

        If (OldPlainTextPassword.Trim = "") OrElse (NewPlainTextPassword.Trim = "") OrElse (ConfirmedPlainTextPassword.Trim = "") Then

            Dim Result As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "All fields must be entered." & vbCrLf & "Please try again.", "Push2Run - Warning", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK)

            If OldPlainTextPassword.Trim.Length = 0 Then
                OldPasswordBox.Focus()
            Else
                If NewPlainTextPassword.Trim.Length = 0 Then
                    WindowNewPasswordBox1.Focus()
                Else
                    WindowNewPasswordBox2.Focus()
                End If
            End If

            Exit Sub

        End If

        ' Check the old password matches the current one on file
        If DoPasswordsMatch(EncryptionClass.Encrypt(OldPlainTextPassword)) Then
        Else
            Dim Result As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "Old password incorrect." & vbCrLf & "Please try again.", "Push2Run - Warning", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK)
            ClearPasswordBoxes()
            Exit Sub
        End If

        If NewPlainTextPassword <> ConfirmedPlainTextPassword Then
            Dim Result As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "New and confirmed passwords don't match." & vbCrLf & "Please try again.", "Push2Run - Warning", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK)
            ClearPasswordBoxes()
            Exit Sub
        End If

        If NewPlainTextPassword = OldPlainTextPassword Then
            Dim Result As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "New and old passwords cannot be the same." & vbCrLf & "Please try again.", "Push2Run - Warning", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK)
            ClearPasswordBoxes()
            Exit Sub
        End If

        'Update the database with the password
        Try

            If SetPassword(NewPlainTextPassword) Then
                gPasswordWasCorrectlyEnteredInPasswordWindow = True
            Else
                Dim Result As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "Password update error." & vbCrLf & "Password not set.", "Push2Run - Warning", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK)
                Result = TopMostMessageBox(gCurrentOwner, "There was a problem updating your password; your database may have been corrupted.", "Push2Run - Warning", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK)
                gPasswordWasCorrectlyEnteredInPasswordWindow = False
            End If

        Catch ex As Exception

            Dim Result As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "Password update error." & vbCrLf & "Password not set.", "Push2Run - Warning", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK)
            gPasswordWasCorrectlyEnteredInPasswordWindow = False

        End Try

        Me.Close()

    End Sub

    <Obfuscation(Feature:="virtualization", Exclude:=False)>
    Private Sub ClearPasswordBoxes()

        OldPasswordBox.Clear()
        WindowNewPasswordBox1.Clear()
        WindowNewPasswordBox2.Clear()

        OldPasswordBox.Focus()

    End Sub

    <Obfuscation(Feature:="virtualization", Exclude:=False)>
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

    <Obfuscation(Feature:="virtualization", Exclude:=False)>
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub ThisWindow_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged

        KeepHelpOnTop()

    End Sub

End Class
