Imports System.ComponentModel

Public Class WindowClosingWarning

    Private Sub WindowClosingWarning_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        SeCursor(CursorState.Normal)
    End Sub

    Private Sub btnOK_Click(sender As Object, e As RoutedEventArgs) Handles btnOK.Click

        Me.Close()

    End Sub

    Private Sub WindowClosingWarning_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

        If cbDoNotShowAgain.IsChecked Then
            My.Settings.ConfirmRedX = False
            My.Settings.Save()
        End If

    End Sub

End Class
