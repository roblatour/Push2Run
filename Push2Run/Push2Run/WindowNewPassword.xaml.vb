Imports System.Reflection

Partial Public Class WindowNewPassword

    
    Private Sub WindowNewPassword_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded

        Me.ShowInTaskbar = False

        MakeTopMost(SafeNativeMethods.FindWindow(Nothing, Me.Title), My.Settings.AlwaysOnTop)

        CheckCapsAndNumLocks()
        PasswordBox1.Focus()

        ResetEncryptionAndDecriptionToReadAndWrite(ResetEncryptionDecriptionLevel.Passwords)

        gPasswordWasCorrectlyEnteredInPasswordWindow = False
        gPasswordWasCorrectlyEnteredInPasswordWindow_UserClickedCancel = False

        SeCursor(CursorState.Normal)

    End Sub

    
    Private Sub WindowChangePassword_Closing(sender As Object, e As ComponentModel.CancelEventArgs) Handles Me.Closing

        ResetEncryptionAndDecriptionToReadAndWrite(ResetEncryptionDecriptionLevel.Data)

    End Sub

    Private Sub Any_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles PasswordBox1.KeyUp, PasswordBox2.KeyUp, Me.KeyUp

        CheckCapsAndNumLocks()

    End Sub

    
    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnOK.Click

        Dim PasswordA As String = PasswordBox1.Password.ToString.Trim
        Dim PasswordB As String = PasswordBox2.Password.ToString.Trim

        If (PasswordA.Length = 0) Then

            Dim Result As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "Password not entered." & vbCrLf & vbCrLf & "Click on the 'OK' button to try again." & vbCrLf & vbCrLf & "Click on the 'Cancel' button if you don't want to set a new password.", "Push2Run - Warning", MessageBoxButton.OKCancel, MessageBoxImage.Exclamation, MessageBoxResult.OK)
            If Result = MessageBoxResult.OK Then
                PasswordBox1.Clear()
                PasswordBox2.Clear()
                PasswordBox1.Focus()
                Exit Sub
            Else
                gPasswordWasCorrectlyEnteredInPasswordWindow = False
                Me.Close()
                Exit Sub
            End If

        End If

        If (PasswordB.Length = 0) Then

            Dim Result As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "Confirmation password not entered." & vbCrLf & vbCrLf & "Click on the 'OK' button to try again." & vbCrLf & vbCrLf & "Click on the 'Cancel' button if you don't want to set a new password.", "Push2Run - Warning", MessageBoxButton.OKCancel, MessageBoxImage.Exclamation, MessageBoxResult.OK)
            If Result = MessageBoxResult.OK Then
                Exit Sub
            Else
                gPasswordWasCorrectlyEnteredInPasswordWindow = False
                Me.Close()
                Exit Sub
            End If

        End If

        If (PasswordA <> PasswordB) Then
            Dim Result As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "Passwords don't match." & vbCrLf & "Please try again.", "Push2Run - Warning", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK)
            PasswordBox1.Clear()
            PasswordBox2.Clear()
            PasswordBox1.Focus()
            Exit Sub
        End If


        'Update the database with the password
        If SetPassword(PasswordA) Then
            gPasswordWasCorrectlyEnteredInPasswordWindow = True
        Else
            Dim Result As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "Password update error." & vbCrLf & "Password not set.", "Push2Run - Warning", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK)
            Result = TopMostMessageBox(gCurrentOwner, "There was a problem updating your password; your database may have been corrupted.", "Push2Run - Warning", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK)
            gPasswordWasCorrectlyEnteredInPasswordWindow = False
        End If

        Me.Close()

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

    
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnCancel.Click
        gPasswordWasCorrectlyEnteredInPasswordWindow = False
        gPasswordWasCorrectlyEnteredInPasswordWindow_UserClickedCancel = True
        Me.Close()
    End Sub

    Private Sub ThisWindow_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged

        KeepHelpOnTop()

    End Sub

End Class

