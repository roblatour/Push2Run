Imports System.ComponentModel
Imports System.IO
Imports System.Net
Imports System.Text
Imports Microsoft.Win32

Public Class WindowUpgradePrompt

    Private lMostCurrentVersion As String = String.Empty
    Private Sub WindowUpgradePrompt_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded

        Try

            gWindowUpgradePromptIsOpen = True

            MakeTopMost(SafeNativeMethods.FindWindow(Nothing, Me.Title), My.Settings.AlwaysOnTop)

            With gCurrentVersionAccordingToWebsite
                lMostCurrentVersion = "v" & .Major & "." & .Minor & "." & .Build & "." & .Revision
            End With

            If My.Settings.SkipUpdateFor = lMostCurrentVersion Then
                cbSkipThisUpdate.IsChecked = True
            End If

            Dim ws As String = New TextRange(RichTextBox1.Document.ContentStart, RichTextBox1.Document.ContentEnd).Text

            Dim sVersionInUse As String = gVersionInUse
            While sVersionInUse.EndsWith(".0")
                sVersionInUse = sVersionInUse.Remove(sVersionInUse.Length - 2)
            End While

            If gBetaVersionInUse Then
                rtbReplace(RichTextBox1, "[Version in use]", sVersionInUse & " beta")
            Else
                rtbReplace(RichTextBox1, "[Version in use]", sVersionInUse)
            End If

            Dim sMostCurrentVersion As String = lMostCurrentVersion
            While sMostCurrentVersion.EndsWith(".0")
                sMostCurrentVersion = sMostCurrentVersion.Remove(sMostCurrentVersion.Length - 2)
            End While
            rtbReplace(RichTextBox1, "[Current version]", sMostCurrentVersion)

            If sVersionInUse = sMostCurrentVersion Then

                If gBetaVersionInUse Then
                Else
                    AutomaticUpdate.Inlines.Clear()
                    AutomaticUpdate.IsEnabled = False
                    cbSkipThisUpdate.Visibility = Visibility.Hidden
                End If

            End If

            WebPageDonateFromUpgrade.NavigateUri = New Uri(gWebPageDonate)

            ' load change log
            Dim myWebClient As System.Net.WebClient = New System.Net.WebClient
            Dim webfilename As String = gWebPageChangeLog

            Dim CurrentDataFileContents As String = myWebClient.DownloadString(webfilename)
            myWebClient.Dispose()

            Dim byteArray As Byte() = Encoding.ASCII.GetBytes(CurrentDataFileContents)
            Dim stream As MemoryStream = New MemoryStream(byteArray)
            rtbChangeLog.Selection.Load(stream, DataFormats.Rtf)

            cbSkipThisUpdate.Content = "Don't remind me about the " & sMostCurrentVersion & " update again"

            SeCursor(CursorState.Normal)

        Catch ex As Exception

            Log(ex.ToString)
            Me.Close()

        End Try

    End Sub
    Private Sub Hyperlink_RequestNavigate(ByVal sender As Object, ByVal e As RequestNavigateEventArgs)
        Process.Start(New ProcessStartInfo(e.Uri.AbsoluteUri))
        e.Handled = True
        Me.Hide()
        System.Threading.Thread.Sleep(1000)
        Me.Close()
    End Sub
    Private Sub AutomaticUpdate_RequestCodeAction(sender As Object, e As RequestNavigateEventArgs)

        AutoUpgrade()
        Me.Close()

    End Sub
    Private Sub AutoUpgrade()

        Try

            ' if the download file for the currently released version already exists then keep it 
            ' otherwise delete any other version

            If File.Exists(gAutomaticUpdateLocalDownloadedFileName) Then

                If IsDownloadFileForTheCurrentlyReleasedVersion() Then
                Else
                    File.Delete(gAutomaticUpdateLocalDownloadedFileName)
                End If

            End If

            'if the download file doesn't exist then download it 

            If File.Exists(gAutomaticUpdateLocalDownloadedFileName) Then

            Else

                Dim OKToTryAutoDownloadAndInstall As Boolean

                Dim Push2RunIsCurrentlnstalledHere As String = AppDomain.CurrentDomain.BaseDirectory.ToUpper

                ' ref: https://stackoverflow.com/questions/23304823/environment-specialfolder-programfiles-returns-the-wrong-directory
                Dim NormalInstallLocation_1, NormalInstallLocation_2 As String
                NormalInstallLocation_1 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86).ToUpper
                NormalInstallLocation_2 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles).ToUpper
                If NormalInstallLocation_1 = NormalInstallLocation_2 Then
                    Try
                        NormalInstallLocation_2 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion")?.GetValue("ProgramFilesDir").ToString.ToUpper
                    Catch ex As Exception
                    End Try
                End If

#If DEBUG Then

                OKToTryAutoDownloadAndInstall = True

#Else

                If Push2RunIsCurrentlnstalledHere.StartsWith(NormalInstallLocation_1) OrElse Push2RunIsCurrentlnstalledHere.StartsWith(NormalInstallLocation_2) Then

                    If HasTheDownloadFileBeenDownloadedInTheLast24Hours() Then
                        OKToTryAutoDownloadAndInstall = False
                    Else
                        OKToTryAutoDownloadAndInstall = True
                    End If

                Else

                    OKToTryAutoDownloadAndInstall = False

                End If

#End If


                If OKToTryAutoDownloadAndInstall Then

                    Dim webClient As WebClient = New WebClient
                    webClient.DownloadFile(gAutomaticUpdateWebFileName, gAutomaticUpdateLocalDownloadedFileName)
                    Log(gAutomaticUpdateLocalDownloadedFileName & " downloaded")
                    Log("")

                    My.Settings.LastDownload = Now
                    My.Settings.Save()

                Else

                    Process.Start(New ProcessStartInfo(gWebPageDownload))
                    Me.Hide()
                    System.Threading.Thread.Sleep(1000)
                    Me.Close()
                    Exit Sub

                End If

            End If

            ' double check everything and start install

            If File.Exists(gAutomaticUpdateLocalDownloadedFileName) Then

                If IsDownloadFileForTheCurrentlyReleasedVersion() Then

                    If ConfirmDownloadFileIsCorrectlySigned(gAutomaticUpdateLocalDownloadedFileName) Then

                        If gIsAdministrator Then

                            Dim TriggerFile As String = Path.GetTempPath & gPush2RunTriggerFileName
                            File.WriteAllText(TriggerFile, "This file is used to trigger Push2Run to start with administrator privileges following an update")

                            'TopMostMessageBox(Me, "Please note" & vbCrLf & vbCrLf &
                            '       "Once Push2Run Is automatically updated it will restart." & vbCrLf & vbCrLf &
                            '       "However, it will no longer be running with Administrator privileges. " & vbCrLf & vbCrLf &
                            '       "If you wish to restore Push2Run's administrator privileges after it has restarted, " &
                            '       "click on:" & vbCrLf & vbCrLf &
                            '       "   'Actions' - 'Give Push2Run administrator privileges'" & vbCrLf & vbCrLf &
                            '       "in Push2Run's Main window's menu.",
                            '       "Push2Run", MessageBoxButton.OK, MessageBoxImage.Exclamation, System.Windows.MessageBoxOptions.None)

                        End If

                        Log("Requesting upgrade")
                        Log("")

                        Dim SilentInstallSettings As String = " /silent /norestart /autoupgrade=1 /log=""Push2RunInstallLog.txt"""

                        Dim Result As ActionStatus = RunProgram(gAutomaticUpdateLocalDownloadedFileName, Path.GetTempPath, SilentInstallSettings, True, ProcessWindowStyle.Normal, "")

                        If Result = ActionStatus.Succeeded Then
                            Log("Automatic update started")
                            Log("")
                        ElseIf Result = ActionStatus.Failed Then
                            Log("Automatic update not started")
                            Log("")
                        End If

                    Else

                        LogAndDislayMessage("The file which was downloaded for an Automatic update can't be used")

                    End If

                Else

                    LogAndDislayMessage("The file which was automatically downloaded wasn't the correct version")

                End If

            End If

        Catch ex As Exception

            Log("Automatic update failed")

        End Try

    End Sub

    Private Sub LogAndDislayMessage(ByVal message As String)

        Log(message)
        Log("")

        Dim Result As MessageBoxResult = TopMostMessageBox(Me, message,
            "Push2Run - Version Check", MessageBoxButton.OK, MessageBoxImage.Information, MessageBoxResult.OK, System.Windows.MessageBoxOptions.None)

    End Sub

    Private Function ConfirmDownloadFileIsCorrectlySigned(ByVal FileName As String) As Boolean

        Dim ReturnValue As Boolean = False

        If gSpecialProcessingForTesting Then

            ReturnValue = True

        Else

            Try

                Dim verify As System.Security.Cryptography.X509Certificates.X509Certificate = System.Security.Cryptography.X509Certificates.X509Certificate.CreateFromSignedFile(FileName)

                If verify.Subject.ToString.StartsWith("CN=Rob Latour,") Then

                    ReturnValue = True

                End If

            Catch ex As Exception

            End Try

        End If

        Return ReturnValue

    End Function

    Private Function HasTheDownloadFileBeenDownloadedInTheLast24Hours() As Boolean

        Dim ReturnValue As Boolean = True

        Try

            If My.Settings.LastDownload = Nothing Then
                ReturnValue = False
            Else
                Dim TwentyFourHoursAgo = Now.AddMilliseconds(-1 * gTwentyFourHoursInMilliSeconds)
                ReturnValue = (My.Settings.LastDownload >= TwentyFourHoursAgo)
            End If

        Catch ex As Exception
        End Try

        Return ReturnValue

    End Function

    Private Sub WindowUpgradePrompt_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

        If cbSkipThisUpdate.IsChecked Then
            My.Settings.SkipUpdateFor = lMostCurrentVersion
        Else
            My.Settings.SkipUpdateFor = "do not skip update"
        End If
        My.Settings.Save()

        gCurrentOwner = ghCurrentOwner
        gWindowUpgradePromptIsOpen = False

    End Sub

End Class