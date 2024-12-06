Imports System.ComponentModel

Public Class WindowAbout
    Private Sub Hyperlink_RequestNavigate(ByVal sender As Object, ByVal e As RequestNavigateEventArgs)
        Process.Start(New ProcessStartInfo(e.Uri.AbsoluteUri))
        e.Handled = True
        Me.Hide()
        System.Threading.Thread.Sleep(1000)
        Me.Close()
    End Sub

    Private Sub WindowAbout_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        RaiseAnEventToShowAboutClosedInSystrayAndMainMenu()
    End Sub

    Private Sub WindowAbout_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded

        Me.Title = gAboutHelpWindowTitle

        MakeTopMost(SafeNativeMethods.FindWindow(Nothing, Me.Title), My.Settings.AlwaysOnTop)

        Dim lVersionInUse As String = gVersionInUse
        While lVersionInUse.EndsWith(".0")
            lVersionInUse = lVersionInUse.Remove(lVersionInUse.Length - 2)
        End While

        If gBetaVersionInUse Then
            rtbReplace(RichTextBox1, "[Version in use]", lVersionInUse & " beta")
        Else
            rtbReplace(RichTextBox1, "[Version in use]", lVersionInUse)
        End If

        WebPageHome.NavigateUri = New Uri(gWebPageHomePage)
        WebPageHelp.NavigateUri = New Uri(gWebPageHelp)
        WebPageLicense.NavigateUri = New Uri(gWebPageLicense)
        WebPageDonate.NavigateUri = New Uri(gWebPageDonate)

        SeCursor(CursorState.Normal)

    End Sub
End Class