Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports System.Text.RegularExpressions

Partial Public Class WindowOptions

    Private WithEvents cbAlwaysOnTop As CheckBox = New CheckBox

    Private WithEvents cbConfirmDelete As CheckBox = New CheckBox
    Private WithEvents cbConfirmExit As CheckBox = New CheckBox
    Private WithEvents cbConfirmRedX As CheckBox = New CheckBox

    Private WithEvents cbCheckForUpdate As CheckBox = New CheckBox
    Private WithEvents rbDaily As RadioButton = New RadioButton
    Private WithEvents rbWeekly As RadioButton = New RadioButton
    Private WithEvents rbEveryTwoWeeks As RadioButton = New RadioButton
    Private WithEvents rbMonthly As RadioButton = New RadioButton
    Private WithEvents btnCheckForUpdateNow As Button = New Button

    Private WithEvents cbRequireAPassword As CheckBox = New CheckBox
    Private WithEvents btnChangeAPassword As Button = New Button
    Private WithEvents pbPushBulletAPI As PasswordBox = New PasswordBox

    Private WithEvents cbStartPush2RunAtLogon As CheckBox = New CheckBox
    Private WithEvents cbStartPush2RunAtLogonAsAdministrator As CheckBox = New CheckBox
    Private WithEvents cbShowPush2RunAtStartup As CheckBox = New CheckBox
    Private WithEvents cbSuppressStartupNotice As CheckBox = New CheckBox

    Private WithEvents cbAutoScroll As CheckBox = New CheckBox
    Private WithEvents cbWriteLogToDisk As CheckBox = New CheckBox

    Private lblHeader As Label = New Label
    Private lblSpacer1 As Label = New Label
    Private lblSpacer2 As Label = New Label
    Private lblSpacer3 As Label = New Label
    Private lblSpacer4 As Label = New Label

    Private WithEvents cbUseDropbox As CheckBox = New CheckBox
    Private lblDropboxLine1 As Label = New Label
    Private WithEvents tbDropboxPath As TextBox = New TextBox
    Private lblDropboxLine3 As Label = New Label
    Private WithEvents tbDropboxFileName As TextBox = New TextBox
    Private lblDropboxLine5 As Label = New Label
    Private WithEvents tbDropboxDeviceName As TextBox = New TextBox

    Private lblConfirm As Label = New Label
    Private WithEvents cbImportTag As CheckBox = New CheckBox
    Private WithEvents cbImportOnByDefault As CheckBox = New CheckBox
    Private WithEvents cbImportConfirmation As CheckBox = New CheckBox

    Private WithEvents cbUseMQTT As CheckBox = New CheckBox
    Private lblMQTTBroker As Label = New Label
    Private lblMQTTPort As Label = New Label
    Private lblMQTTUser As Label = New Label
    Private lblMQTTPassword As Label = New Label
    Private lblMQTTFilter As Label = New Label
    Private WithEvents tbMQTTBroker As TextBox = New TextBox
    Private WithEvents tbMQTTPort As TextBox = New TextBox
    Private WithEvents tbMQTTUser As TextBox = New TextBox
    Private WithEvents pbMQTTPassword As PasswordBox = New PasswordBox
    Private WithEvents tbMQTTFilter As TextBox = New TextBox
    Private WithEvents cbMQTTListenForPayloadOnly As CheckBox = New CheckBox

    Private WithEvents cbShowNotifications As CheckBox = New CheckBox
    Private WithEvents cbShowNotificationResult As CheckBox = New CheckBox
    Private WithEvents cbShowNotificationSource As CheckBox = New CheckBox
    Private WithEvents cbIncludeDisconnectAndReconnect As CheckBox = New CheckBox
    Private WithEvents btnTestNotification As Button = New Button

    Private WithEvents cbUsePushbullet As CheckBox = New CheckBox
    Private lblPushbulletLine1 As Label = New Label
    Private lblPushbulletLine3 As Label = New Label
    Private WithEvents tbPushBulletTitleFilter As TextBox = New TextBox

    Private WithEvents cbUsePushover As CheckBox = New CheckBox
    Private lblPushoverLine1 As Label = New Label
    Private WithEvents tbPushoverUserID As TextBox = New TextBox
    Private lblPushoverLine3 As Label = New Label
    Private WithEvents tbPushoverDeviceName As TextBox = New TextBox
    Private lblPushoverLine5 As Label = New Label
    Private WithEvents btnAuthenticate As Button = New Button

    Private lblSeparatingWordsLine1 As Label = New Label
    Private WithEvents tbSeparatingWords As TextBox = New TextBox

    Private lblDatabaseLocationLine4 As Label = New Label
    Private WithEvents btnDatabaseLocationLine5 As Button = New Button
    Private lblDatabaseLocationLine3 As Label = New Label
    Private lblDatabaseLocationLine1 As Label = New Label
    Private WithEvents btnDatabaseLocationLine2 As Button = New Button
    Private lblDatabaseLocationLine6 As Label = New Label
    Private lblDatabaseLocationLine7 As Label = New Label

    Private lblLogFileLocationLine1 As Label = New Label
    Private WithEvents btnLogFileLocationLine2 As Button = New Button
    Private lblLogFileLocationLine3 As Label = New Label
    Private lblLogFileLocationLine4 As Label = New Label

    Private WithEvents cbUACLimit As CheckBox = New CheckBox

    Private LoadUnderway As Boolean = True

    Private OriginalPasswordRequired As Boolean = False
    Private OriginalEncryptedPassword As String = String.Empty

    Const IdealWidth As Double = 435
    Const IdealHeight As Double = 22


    Private Sub WindowOptions_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded

        Me.ShowInTaskbar = False

        MakeTopMost(SafeNativeMethods.FindWindow(Nothing, Me.Title), My.Settings.AlwaysOnTop)

        '************************************************************************************

        'v4.6
        lblSpacer1.IsTabStop = False
        lblSpacer2.IsTabStop = False
        lblSpacer3.IsTabStop = False
        lblSpacer4.IsTabStop = False

        '************************************************************************************

        cbAlwaysOnTop.IsChecked = My.Settings.AlwaysOnTop
        cbAlwaysOnTop.Content = "_Keep Push2Run windows on top of other windows"

        '************************************************************************************

        lblConfirm.Content = "Confirm when:"

        cbConfirmDelete.IsChecked = My.Settings.ConfirmDelete
        cbConfirmDelete.Content = "_Deleting"
        cbConfirmDelete.Margin = New Thickness(12, 0, 0, 0)


        cbConfirmExit.IsChecked = My.Settings.ConfirmExit
        cbConfirmExit.Content = "_Exiting"
        cbConfirmExit.Margin = New Thickness(12, 0, 0, 0)

        cbConfirmRedX.IsChecked = My.Settings.ConfirmRedX
        cbConfirmRedX.Content = "_Hiding the main window"
        cbConfirmRedX.Margin = New Thickness(12, 0, 0, 0)

        cbImportConfirmation.IsChecked = My.Settings.ImportConfirmation
        cbImportConfirmation.Content = "_Importing is done"
        cbImportConfirmation.Margin = New Thickness(12, 0, 0, 0)

        '************************************************************************************

        cbAutoScroll.IsChecked = My.Settings.AutoScroll
        cbAutoScroll.Content = "_Auto scroll the Session Log"

        cbWriteLogToDisk.IsChecked = My.Settings.WriteLogToDisk
        cbWriteLogToDisk.Content = "_Write session log to disk"

        lblLogFileLocationLine1.Content = "Session _log file:"
        lblLogFileLocationLine1.Margin = New Thickness(12, 0, 0, 0)
        btnLogFileLocationLine2.Content = gSessionLogFile

        btnLogFileLocationLine2.FontWeight = System.Windows.FontWeights.SemiBold
        btnLogFileLocationLine2.Margin = New Thickness(16, 0, 0, 0)

        lblLogFileLocationLine3.Content = " "
        lblLogFileLocationLine4.Content = "The Session log file location and name may not be changed."


        '************************************************************************************
        cbCheckForUpdate.IsChecked = My.Settings.CheckForUpdate
        cbCheckForUpdate.Content = "_Automatically check for updates:"
        rbDaily.Content = "_Daily"
        rbWeekly.Content = "_Weekly"
        rbEveryTwoWeeks.Content = "_Every two weeks"
        rbMonthly.Content = "_Monthly"
        rbDaily.Margin = New Thickness(17, 0, 0, 0) 'indent the radio buttons
        rbWeekly.Margin = New Thickness(17, 0, 0, 0)
        rbEveryTwoWeeks.Margin = New Thickness(17, 0, 0, 0)
        rbMonthly.Margin = New Thickness(17, 0, 0, 0)
        btnCheckForUpdateNow.Content = "_Check for update now"
        btnCheckForUpdateNow.Width = IdealWidth
        btnAuthenticate.Content = "_Authenticate"
        btnAuthenticate.Width = IdealWidth

        rbDaily.Tag = "1"
        rbWeekly.Tag = "2"
        rbEveryTwoWeeks.Tag = "3"
        rbMonthly.Tag = "4"

        SetLookOfCheckForUpdate(My.Settings.CheckForUpdateFrequency)

        '************************************************************************************

        lblDatabaseLocationLine1.Content = "Database file:"

        btnDatabaseLocationLine2.Content = gSQLiteFullDatabaseName
        btnDatabaseLocationLine2.FontWeight = System.Windows.FontWeights.SemiBold
        btnDatabaseLocationLine2.Margin = New Thickness(5, 0, 0, 0)
        btnDatabaseLocationLine2.Width = IdealWidth
        btnDatabaseLocationLine2.ToolTip = "Click to view in File Explore"

        lblDatabaseLocationLine3.Content = " "
        lblDatabaseLocationLine4.Content = "Settings file:"

        btnDatabaseLocationLine5.Content = Configuration.ConfigurationManager.OpenExeConfiguration(System.Configuration.ConfigurationUserLevel.PerUserRoamingAndLocal).FilePath.Replace("_", "__")
        btnDatabaseLocationLine5.FontWeight = System.Windows.FontWeights.SemiBold
        btnDatabaseLocationLine5.Margin = New Thickness(5, 0, 0, 0)
        btnDatabaseLocationLine5.Width = IdealWidth
        btnDatabaseLocationLine5.ToolTip = "Click to view in File Explore"

        lblDatabaseLocationLine6.Content = " "
        lblDatabaseLocationLine7.Content = "These file locations and names may not be changed."

        '************************************************************************************
        cbRequireAPassword.IsChecked = IsAPasswordRequiredForBoss()
        cbRequireAPassword.Content = "_Password required"
        btnChangeAPassword.Content = "_Change password"

        '************************************************************************************
        cbStartPush2RunAtLogon.IsChecked = My.Settings.StartBossOnLogon
        cbStartPush2RunAtLogon.Content = "Automatically start Push2Run when you log on to Windows"

        cbStartPush2RunAtLogonAsAdministrator.IsChecked = My.Settings.StartBossAsAdministratorByDefault
        cbStartPush2RunAtLogonAsAdministrator.Content = "When automatically starting at Windows log on use administrator privileges"

        If cbStartPush2RunAtLogon.IsChecked Then
        Else
            cbStartPush2RunAtLogonAsAdministrator.IsChecked = False
            cbStartPush2RunAtLogonAsAdministrator.IsEnabled = False
            cbStartPush2RunAtLogonAsAdministrator.Foreground = Brushes.Gray
        End If

        cbShowPush2RunAtStartup.IsChecked = My.Settings.ShowPush2RunAtStartup
        cbShowPush2RunAtStartup.Content = "Show the main _window at start-up"

        cbSuppressStartupNotice.IsChecked = My.Settings.SuppressStartupNotice
        cbSuppressStartupNotice.Content = "Suppress start-up and other _notices when no triggers are enabled"

        '************************************************************************************
        lblSpacer1.Content = " "
        lblSpacer2.Content = " "
        lblSpacer3.Content = " "
        lblSpacer4.Content = " "
        lblSpacer1.Height = 10
        lblSpacer2.Height = 10
        lblSpacer3.Height = 10
        lblSpacer4.Height = 10


        '************************************************************************************
        ' lblHeader.Content = "Dropbox:"
        cbUseDropbox.IsChecked = My.Settings.UseDropbox
        cbUseDropbox.Content = "Enable _Dropbox"
        lblHeader.FontWeight = System.Windows.FontWeights.SemiBold
        lblDropboxLine1.Content = "Dropbox folder path:"
        tbDropboxPath.Text = My.Settings.DropboxPath.Trim
        tbDropboxPath.Width = IdealWidth
        lblDropboxLine3.Content = "Dropbox file name:"
        tbDropboxFileName.Text = My.Settings.DropboxFileName.Trim
        tbDropboxFileName.Width = IdealWidth
        lblDropboxLine5.Content = "Device name:"
        tbDropboxDeviceName.Text = My.Settings.DropboxDeviceName.Trim
        tbDropboxDeviceName.Width = IdealWidth

        '************************************************************************************
        'lblHeader.Content = "Imports:"
        cbImportOnByDefault.IsChecked = My.Settings.ImportOnByDefault
        cbImportOnByDefault.Content = "_Automatically turn on imported cards"

        cbImportTag.IsChecked = My.Settings.ImportTag
        cbImportTag.Content = "_Update the description of imported cards to end with '(Imported)'"

        '************************************************************************************
        'lblHeader.Content = "Notifications:"
        cbShowNotifications.IsChecked = My.Settings.ShowNotifications
        cbShowNotifications.Content = "_Show notifications when Push2Run receives a command"

        cbShowNotificationResult.IsChecked = My.Settings.ShowNotificationResult
        cbShowNotificationResult.Content = "Include _results"
        cbShowNotificationResult.Margin = New Thickness(17, 0, 0, 0)

        cbShowNotificationSource.IsChecked = My.Settings.ShowNotificationSource
        cbShowNotificationSource.Content = "Include _triggers"
        cbShowNotificationSource.Margin = New Thickness(17, 0, 0, 0)

        btnTestNotification.Content = "_Test"
        btnTestNotification.FontWeight = System.Windows.FontWeights.SemiBold
        btnTestNotification.Margin = New Thickness(17, 0, 0, 0)
        btnTestNotification.Width = IdealWidth - btnTestNotification.Margin.Left

        cbIncludeDisconnectAndReconnect.IsChecked = My.Settings.IncludeDisconnectAndReconnect
        cbIncludeDisconnectAndReconnect.Content = "Show notifications _if MQTT/Pushbullet/Pushover disconnects or reconnects"

        '************************************************************************************
        'lblHeader.Content = "MQTT:"
        lblHeader.FontWeight = System.Windows.FontWeights.SemiBold

        cbUseMQTT.IsChecked = My.Settings.UseMQTT

        cbUseMQTT.Content = "Enable _MQTT"
        lblMQTTBroker.Content = "_Broker:"
        lblMQTTPort.Content = "_Port:"
        lblMQTTUser.Content = "_User:"
        lblMQTTPassword.Content = "Pass_word:"
        lblMQTTFilter.Content = "_Topic filter(s):"
        cbMQTTListenForPayloadOnly.Content = "Listen for Payload only"
        cbMQTTListenForPayloadOnly.ToolTip = "check to listen for on Payload only, uncheck to listen for Topic/Payload"
        ToolTipService.SetInitialShowDelay(cbMQTTListenForPayloadOnly, 200)
        ToolTipService.SetShowDuration(cbMQTTListenForPayloadOnly, 10000)

        tbMQTTBroker.Text = My.Settings.MQTTBroker
        tbMQTTPort.Text = My.Settings.MQTTPort
        tbMQTTUser.Text = My.Settings.MQTTUser
        pbMQTTPassword.Password = EncryptionClass.Decrypt(My.Settings.MQTTPassword)
        tbMQTTFilter.Text = My.Settings.MQTTFilter
        cbMQTTListenForPayloadOnly.IsChecked = My.Settings.MQTTListenForPayloadOnly

        lblMQTTBroker.Width = 80
        lblMQTTPort.Width = 80
        lblMQTTUser.Width = 80
        lblMQTTPassword.Width = 80
        lblMQTTFilter.Width = 80

        tbMQTTBroker.Width = 100
        tbMQTTPort.Width = 100
        tbMQTTUser.Width = 150
        pbMQTTPassword.Width = 150
        tbMQTTFilter.Width = 365

        tbMQTTBroker.Height = IdealHeight
        tbMQTTPort.Height = IdealHeight
        tbMQTTUser.Height = IdealHeight
        pbMQTTPassword.Height = IdealHeight
        tbMQTTFilter.Height = IdealHeight '* 2
        'tbMQTTFilter.TextWrapping = TextWrapping.Wrap

        cbMQTTListenForPayloadOnly.Height = IdealHeight

        tbMQTTBroker.VerticalContentAlignment = VerticalAlignment.Center
        tbMQTTPort.VerticalContentAlignment = VerticalAlignment.Center
        tbMQTTUser.VerticalContentAlignment = VerticalAlignment.Center
        pbMQTTPassword.VerticalContentAlignment = VerticalAlignment.Center
        tbMQTTFilter.VerticalContentAlignment = VerticalAlignment.Top

        '************************************************************************************
        'lblHeader.Content = "Pushbullet:"
        lblHeader.FontWeight = System.Windows.FontWeights.SemiBold

        cbUsePushbullet.IsChecked = My.Settings.UsePushbullet
        cbUsePushbullet.Content = "_Enable Pushbullet:"
        lblPushbulletLine1.Content = "Pushbullet Access _Token:"
        pbPushBulletAPI.Password = EncryptionClass.Decrypt(My.Settings.PushBulletAPI)
        pbPushBulletAPI.Width = IdealWidth
        lblPushbulletLine3.Content = "Title _filter"
        tbPushBulletTitleFilter.Text = My.Settings.PushBulletTitleFilter
        tbPushBulletTitleFilter.Width = IdealWidth

        '************************************************************************************
        ' lblHeader.Content = "Pushover:"
        cbUsePushover.IsChecked = My.Settings.UsePushover
        cbUsePushover.Content = "_Enable Pushover"
        lblHeader.FontWeight = System.Windows.FontWeights.SemiBold
        lblPushoverLine1.Content = "Pushover _user id:"
        tbPushoverUserID.Text = My.Settings.PushoverUserID
        tbPushoverUserID.Width = IdealWidth
        lblPushoverLine3.Content = "Pushover _device name:"
        tbPushoverDeviceName.Text = My.Settings.PushoverDeviceName.Trim
        tbPushoverDeviceName.Width = IdealWidth
        lblPushoverLine5.Content = ""

        '************************************************************************************
        ' lblHeader.Content = "Separating words:"
        lblSeparatingWordsLine1.Content = "_Separating words:"
        tbSeparatingWords.Text = My.Settings.SeparatingWords
        tbSeparatingWords.Width = IdealWidth

        '************************************************************************************

        cbUACLimit.IsChecked = My.Settings.UACLimit
        cbUACLimit.Content = "_Don't run a program if a UAC prompt is required"

        '************************************************************************************

        StackPanel1.Children.Add(cbAlwaysOnTop)

        ScrollViewer1.Content = StackPanel1
        ScrollViewer1.Height = 200

        OriginalPasswordRequired = IsAPasswordRequiredForBoss()
        If OriginalPasswordRequired Then
            OriginalEncryptedPassword = GetEncryptedPassword()
        End If

        Me.AlwaysOnTop.Focus()

        Select Case gOpenOptionsWindowAt

            Case Is = gOpenOptions.Dropbox
                SetSelectedTreeViewItem("Dropbox")

            Case Is = gOpenOptions.Pushbullet
                SetSelectedTreeViewItem("Pushbullet")

            Case Is = gOpenOptions.Pushover
                SetSelectedTreeViewItem("Pushover")

            Case Else
                SetSelectedTreeViewItem("Always on top")

        End Select

        gOpenOptionsWindowAt = gOpenOptions.AlwaysOnTop 'reset for next time

        If (tbDropboxDeviceName.Text = String.Empty) OrElse (tbDropboxDeviceName.Text = gNotAvailable) Then
            tbDropboxDeviceName.Text = "Push2Run_" & My.Computer.Name.Trim
        End If

        If (tbDropboxPath.Text = String.Empty) OrElse (tbDropboxPath.Text = gNotAvailable) Then
            tbDropboxPath.Text = Environ$("USERPROFILE") & gDefaultDropboxPath
        End If

        If (tbPushBulletTitleFilter.Text = String.Empty) OrElse (tbPushBulletTitleFilter.Text = gNotAvailable) Then
            tbPushBulletTitleFilter.Text = "Push2Run " & My.Computer.Name.Trim
        End If

        If (tbPushoverDeviceName.Text = String.Empty) OrElse (tbPushoverDeviceName.Text = gNotAvailable) Then
            tbPushoverDeviceName.Text = "Push2Run_" & My.Computer.Name.Trim
        End If

        AddHandler OptionsClosed, AddressOf CloseMeNow

        SeCursor(CursorState.Normal)

        LoadUnderway = False

    End Sub

    Private Sub CloseMeNow()

        Me.Close()

    End Sub

    Private Sub SetSelectedTreeViewItem(ByVal search As String)

        Try

            For Each item As TreeViewItem In TreeView1.ItemContainerGenerator.Items

                If item.Header.trim = search Then
                    item.IsSelected = True
                    Exit For
                End If

            Next

        Catch ex As Exception
        End Try

    End Sub

    Private Sub tbPushBulletTitleFilter_MouseLeave(sender As Object, e As RoutedEventArgs) Handles tbPushBulletTitleFilter.MouseLeave

        tbPushBulletTitleFilter.Text = CleanUpWhiteAndDuplicatedSpaces(tbPushBulletTitleFilter.Text)
        tbPushBulletTitleFilter.Select(tbPushBulletTitleFilter.Text.Length, 0)

    End Sub


    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnOK.Click

        tbPushBulletTitleFilter.Text = CleanUpWhiteAndDuplicatedSpaces(tbPushBulletTitleFilter.Text)

        gCurrentOwner = Me

        If cbUseDropbox.IsChecked Then

            If tbDropboxFileName.Text.Trim.Length = 0 Then

                Beep()

                If TopMostMessageBox(gCurrentOwner, "The Dropbox filename cannot be blank." & vbCrLf & vbCrLf &
                               "Click 'OK' to use the default value which is" & vbCrLf & gDefaultDropboxFilename & vbCrLf & "or 'Cancel' to try again.",
                               "Push2Run - Warning", MessageBoxButton.OKCancel, MessageBoxImage.Exclamation, MessageBoxResult.OK) = MessageBoxResult.OK Then

                    tbDropboxFileName.Text = gDefaultDropboxFilename

                Else

                    SetSelectedTreeViewItem("Dropbox")
                    Exit Sub

                End If

            End If

            If IsValidFileName(tbDropboxFileName.Text.Trim) Then
            Else

                Beep()

                If TopMostMessageBox(gCurrentOwner, "The Dropbox filename is invalid." & vbCrLf & vbCrLf &
                               "Click 'OK' to use the default value which is" & vbCrLf & gDefaultDropboxFilename & vbCrLf & "or 'Cancel' to try again.",
                               "Push2Run - Warning", MessageBoxButton.OKCancel, MessageBoxImage.Exclamation, MessageBoxResult.OK) = MessageBoxResult.OK Then

                    tbDropboxFileName.Text = gDefaultDropboxFilename

                Else

                    SetSelectedTreeViewItem("Dropbox")
                    Exit Sub

                End If

            End If

            If tbDropboxPath.Text.Trim.Length = 0 Then

                Beep()

                Dim defdir As String = Environ$("USERPROFILE") & gDefaultDropboxPath

                If TopMostMessageBox(gCurrentOwner, "The Dropbox path cannot be blank." & vbCrLf & vbCrLf &
                              "Click 'OK' to use the default value, which is " & vbCrLf & defdir & vbCrLf & "or 'Cancel' to try again.",
                              "Push2Run - Warning", MessageBoxButton.OKCancel, MessageBoxImage.Exclamation, MessageBoxResult.OK) = MessageBoxResult.OK Then
                    tbDropboxPath.Text = defdir

                Else

                    SetSelectedTreeViewItem("Dropbox")
                    Exit Sub

                End If

            End If

            tbDropboxPath.Text = tbDropboxPath.Text.Trim

            tbDropboxPath.Text = tbDropboxPath.Text.Replace("/", "\")

            If tbDropboxPath.Text.EndsWith("\") Then
            Else
                tbDropboxPath.Text = tbDropboxPath.Text & "\"
            End If

            If IsValidPathName(tbDropboxPath.Text.Trim) Then

            Else

                Beep()

                Dim defdir As String = Environ$("USERPROFILE") & gDefaultDropboxPath

                If TopMostMessageBox(gCurrentOwner, "The Dropbox path is invalid." & vbCrLf & vbCrLf &
                              "Click 'OK' to use the default value, which is " & vbCrLf & defdir & vbCrLf & "or 'Cancel' to try again.",
                              "Push2Run - Warning", MessageBoxButton.OKCancel, MessageBoxImage.Exclamation, MessageBoxResult.OK) = MessageBoxResult.OK Then
                    tbDropboxPath.Text = defdir
                Else

                    SetSelectedTreeViewItem("Dropbox")
                    Exit Sub
                End If

            End If

            If My.Computer.FileSystem.DirectoryExists(tbDropboxPath.Text.Trim) Then
            Else

                Beep()

                If TopMostMessageBox(gCurrentOwner, "The Dropbox path does not exist." & vbCrLf & vbCrLf &
                               "Click 'OK' to have Push2Run attempt to create it, or 'Cancel' to try again.",
                               "Push2Run - Warning", MessageBoxButton.OKCancel, MessageBoxImage.Exclamation, MessageBoxResult.OK) = MessageBoxResult.OK Then

                    Try

                        Directory.CreateDirectory(tbDropboxPath.Text.Trim)

                        If My.Computer.FileSystem.DirectoryExists(tbDropboxPath.Text.Trim) Then
                            Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "Dropbox path created successfully.", "Push2Run - Info", MessageBoxButton.OK, MessageBoxImage.Exclamation)
                        Else
                            Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "Dropbox path was not created.", "Push2Run - Info", MessageBoxButton.OK, MessageBoxImage.Exclamation)
                            SetSelectedTreeViewItem("Dropbox")
                            Exit Sub
                        End If

                    Catch ex As Exception

                        Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "Dropbox path was not created.", "Push2Run - Info", MessageBoxButton.OK, MessageBoxImage.Exclamation)
                        SetSelectedTreeViewItem("Dropbox")
                        Exit Sub

                    End Try

                Else

                    SetSelectedTreeViewItem("Dropbox")
                    Exit Sub

                End If

            End If

            If (tbDropboxDeviceName.Text = String.Empty) Then

                Beep()

                If TopMostMessageBox(gCurrentOwner, "The Dropbox device name cannot be blank." & vbCrLf & vbCrLf &
                               "Click 'OK' to use the default value for this computer, which is" & vbCrLf & "Push2Run_" & My.Computer.Name & vbCrLf & "or 'Cancel' to try again.",
                               "Push2Run - Warning", MessageBoxButton.OKCancel, MessageBoxImage.Exclamation, MessageBoxResult.OK) = MessageBoxResult.OK Then

                    tbDropboxDeviceName.Text = "Push2Run_" & My.Computer.Name

                Else

                    SetSelectedTreeViewItem("Dropbox")
                    Exit Sub

                End If

            End If

        End If

        If cbUsePushbullet.IsChecked Then

            If tbPushBulletTitleFilter.Text = String.Empty Then

                Beep()

                If TopMostMessageBox(gCurrentOwner, "The Pushbullet Title Filter cannot be blank." & vbCrLf & vbCrLf &
                                     "Click 'OK' to use the default value for this computer, which is" & vbCrLf & "Push2Run " & My.Computer.Name & vbCrLf & "or 'Cancel' to try again.",
                                   "Push2Run - Warning", MessageBoxButton.OKCancel, MessageBoxImage.Exclamation, MessageBoxResult.OK) = MessageBoxResult.OK Then

                    tbPushBulletTitleFilter.Text = "Push2Run " & My.Computer.Name

                Else

                    SetSelectedTreeViewItem("Pushbullet")
                    Exit Sub

                End If

            End If

        End If


        If cbUsePushbullet.IsChecked OrElse cbUsePushover.IsChecked OrElse cbUseDropbox.IsChecked OrElse cbUseMQTT.IsChecked OrElse cbSuppressStartupNotice.IsChecked Then
        Else
            Beep()

            If TopMostMessageBox(gCurrentOwner, "At least one of Dropbox, Pushbullet, Pushover, or MQTT should be enabled for use, currently none are." & vbCrLf & vbCrLf &
                               "Do you want to correct this?",
                               "Push2Run - Warning", MessageBoxButton.YesNo, MessageBoxImage.Exclamation, MessageBoxResult.OK) = MessageBoxResult.Yes Then

                SetSelectedTreeViewItem("Dropbox")
                Exit Sub

            End If

        End If


        If cbUseMQTT.IsChecked Then

            If tbMQTTFilter.Text.Trim.Length = 0 Then

                Beep()

                Dim Dummy = TopMostMessageBox(gCurrentOwner, "IF MQTTT is enabled the MQTT Topic Filter(s) may not be empty.",
                        "Push2Run - Warning", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK)

                SetSelectedTreeViewItem("MQTT")

                Exit Sub

            End If

            Dim Topics() As String = tbMQTTFilter.Text.Trim.Split(" ")

            tbMQTTFilter.Text = String.Empty

            For x = 0 To Topics.Count - 1

                If Topics(x).Trim.Length > 0 Then
                    tbMQTTFilter.Text &= Topics(x).Trim & " "
                End If

            Next

            tbMQTTFilter.Text = tbMQTTFilter.Text.Trim

            If Topics.Length <> Topics.Distinct().Count() Then

                Beep()

                Dim Dummy = TopMostMessageBox(gCurrentOwner, "The MQTT Topic Filter(s) field contains duplicate topics." & vbCrLf & vbCrLf & "Duplicate topics are not allowed.",
                 "Push2Run - Warning", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK)
                SetSelectedTreeViewItem("MQTT")
                Exit Sub

            End If

        End If


        Dim ProblemWithPortNumber As Boolean = False


        Try

            Dim test = CInt(tbMQTTPort.Text)

            If test < 1024 OrElse test > 65535 Then
                ProblemWithPortNumber = True
            End If

        Catch ex As Exception
            ProblemWithPortNumber = True
        End Try


        If ProblemWithPortNumber Then

            If TopMostMessageBox(gCurrentOwner, "Invalid MQTT Port number. The acceptable range for the MQTT Port number is between 1024 and 65535 inclusive, with the default being 1883.  Would you like to use the default?",
                               "Push2Run - Warning", MessageBoxButton.YesNo, MessageBoxImage.Exclamation, MessageBoxResult.OK) = MessageBoxResult.Yes Then
                tbMQTTPort.Text = "1883"
            Else
                Exit Sub
            End If

        End If

        'clean up Separating words
        tbSeparatingWords.Text = tbSeparatingWords.Text.ToLower
        Dim SepWords() As String = tbSeparatingWords.Text.Split(",")
        For x As Integer = 0 To SepWords.Count - 1
            SepWords(x) = SepWords(x).Trim
        Next
        Array.Sort(SepWords)
        Dim SepWordNoDuplicates As HashSet(Of String) = New HashSet(Of String)(SepWords)
        Dim ws As String = String.Empty
        For Each SepWord In SepWordNoDuplicates
            If SepWord.Trim.Length > 0 Then
                ws = ws & ", " & SepWord.Trim
            End If
        Next
        If ws.StartsWith(",") Then
            ws = ws.Remove(0, 1)
        End If
        tbSeparatingWords.Text = ws.Trim

        My.Settings.AlwaysOnTop = cbAlwaysOnTop.IsChecked

        My.Settings.AutoScroll = cbAutoScroll.IsChecked
        My.Settings.WriteLogToDisk = cbWriteLogToDisk.IsChecked

        My.Settings.ConfirmDelete = cbConfirmDelete.IsChecked
        My.Settings.ConfirmExit = cbConfirmExit.IsChecked
        My.Settings.ConfirmRedX = cbConfirmRedX.IsChecked

        My.Settings.CheckForUpdate = cbCheckForUpdate.IsChecked
        If rbDaily.IsChecked Then
            My.Settings.CheckForUpdateFrequency = UpdateCheckFrequency.Daily

        ElseIf rbWeekly.IsChecked Then
            My.Settings.CheckForUpdateFrequency = UpdateCheckFrequency.Weekly

        ElseIf rbEveryTwoWeeks.IsChecked Then
            My.Settings.CheckForUpdateFrequency = UpdateCheckFrequency.EveryTwoWeeks

        Else
            My.Settings.CheckForUpdateFrequency = UpdateCheckFrequency.Monthly
        End If

        My.Settings.StartBossOnLogon = cbStartPush2RunAtLogon.IsChecked
        My.Settings.StartBossAsAdministratorByDefault = cbStartPush2RunAtLogonAsAdministrator.IsChecked
        My.Settings.ShowPush2RunAtStartup = cbShowPush2RunAtStartup.IsChecked
        My.Settings.SuppressStartupNotice = cbSuppressStartupNotice.IsChecked

        My.Settings.UseDropbox = cbUseDropbox.IsChecked
        My.Settings.DropboxPath = tbDropboxPath.Text.Trim
        My.Settings.DropboxFileName = tbDropboxFileName.Text.Trim
        My.Settings.DropboxDeviceName = tbDropboxDeviceName.Text.Trim

        My.Settings.ImportConfirmation = cbImportConfirmation.IsChecked
        My.Settings.ImportOnByDefault = cbImportOnByDefault.IsChecked
        My.Settings.ImportTag = cbImportTag.IsChecked

        My.Settings.UseMQTT = cbUseMQTT.IsChecked
        My.Settings.MQTTBroker = tbMQTTBroker.Text
        My.Settings.MQTTPort = tbMQTTPort.Text
        My.Settings.MQTTUser = tbMQTTUser.Text
        My.Settings.MQTTPassword = EncryptionClass.Encrypt(pbMQTTPassword.Password.ToString.Trim)
        My.Settings.MQTTListenForPayloadOnly = cbMQTTListenForPayloadOnly.IsChecked

        tbMQTTFilter.Text = tbMQTTFilter.Text.Replace(Chr(10), " ").Replace(Chr(130), " ").Trim

        While tbMQTTFilter.Text.Contains("  ")
            tbMQTTFilter.Text = tbMQTTFilter.Text.Replace("  ", " ")
        End While

        My.Settings.MQTTFilter = tbMQTTFilter.Text

        My.Settings.ShowNotifications = cbShowNotifications.IsChecked
        My.Settings.ShowNotificationResult = cbShowNotificationResult.IsChecked
        My.Settings.ShowNotificationSource = cbShowNotificationSource.IsChecked
        My.Settings.IncludeDisconnectAndReconnect = cbIncludeDisconnectAndReconnect.IsChecked

        My.Settings.UsePushbullet = cbUsePushbullet.IsChecked
        My.Settings.PushBulletAPI = EncryptionClass.Encrypt(pbPushBulletAPI.Password.ToString.Trim)
        My.Settings.PushBulletTitleFilter = tbPushBulletTitleFilter.Text.Trim

        My.Settings.UsePushover = cbUsePushover.IsChecked
        My.Settings.PushoverUserID = tbPushoverUserID.Text.Trim
        My.Settings.PushoverDeviceName = tbPushoverDeviceName.Text.Trim

        My.Settings.SeparatingWords = tbSeparatingWords.Text.Trim

        My.Settings.Save()
        Me.Close()

        'Update Global Separating words array for later use

        UpdateGlobalSeparatingWordsArray()

    End Sub


    Private Sub BtnCancel_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles BtnCancel.Click

        gEnteredPassword = String.Empty

        If OriginalPasswordRequired Then

            If IsAPasswordRequiredForBoss() Then

                Dim CurrentEncryptedPassword As String = GetEncryptedPassword()
                If CurrentEncryptedPassword <> OriginalEncryptedPassword Then
                    ResetEncryptionAndDecriptionToReadAndWrite(ResetEncryptionDecriptionLevel.Passwords)
                    SetPassword(EncryptionClass.Decrypt(OriginalEncryptedPassword))
                End If

            Else

                SetPasswordIsRequireFlag(True)
                ResetEncryptionAndDecriptionToReadAndWrite(ResetEncryptionDecriptionLevel.Passwords)
                SetPassword(EncryptionClass.Decrypt(OriginalEncryptedPassword))

            End If

        Else

            If IsAPasswordRequiredForBoss() Then
                SetPasswordIsRequireFlag(False)
                SetPassword(String.Empty)
            Else
                ' do nothing
            End If

        End If

        StartupShortCut("Ensure")

        Me.Close()

    End Sub


    Private Sub SetLookOfCheckForUpdate(ByVal SelectionChoice As Integer)

        rbDaily.IsEnabled = cbCheckForUpdate.IsChecked
        rbWeekly.IsEnabled = cbCheckForUpdate.IsChecked
        rbEveryTwoWeeks.IsEnabled = cbCheckForUpdate.IsChecked
        rbMonthly.IsEnabled = cbCheckForUpdate.IsChecked

        Select Case SelectionChoice
            Case UpdateCheckFrequency.Daily
                rbDaily.IsChecked = True
            Case UpdateCheckFrequency.Weekly
                rbWeekly.IsChecked = True
            Case UpdateCheckFrequency.EveryTwoWeeks
                rbEveryTwoWeeks.IsChecked = True
            Case UpdateCheckFrequency.Monthly
                rbMonthly.IsChecked = True
        End Select

    End Sub


    Private Sub TreeView1_SelectedItemChanged(ByVal sender As Object, ByVal e As System.Windows.RoutedPropertyChangedEventArgs(Of Object)) Handles TreeView1.SelectedItemChanged

        BuildStackPanel(sender.selecteditem.tag)

    End Sub


    Private Sub BuildStackPanel(ByVal SpecificPanel As String)

        ' The WorkingPanel1x variables, used for MQTT opton, need to be defined like this to avoid the error
        ' Specified element is already the logical child of another element. Disconnect it first.'
        ' when this routine is called multiple times

        Static WorkingPanel1 As StackPanel = New StackPanel()
        Static WorkingPanel2 As StackPanel = New StackPanel()
        Static WorkingPanel3 As StackPanel = New StackPanel()
        Static WorkingPanel4 As StackPanel = New StackPanel()
        Static WorkingPanel5 As StackPanel = New StackPanel()
        Static WorkingPanel6 As StackPanel = New StackPanel()

        StackPanel1.Children.Clear()

        Select Case SpecificPanel

            Case Is = "Always on top"
                StackPanel1.Children.Add(cbAlwaysOnTop)

            Case Is = "Confirmations"
                StackPanel1.Children.Add(lblConfirm)
                StackPanel1.Children.Add(cbConfirmDelete)
                StackPanel1.Children.Add(lblSpacer1)
                StackPanel1.Children.Add(cbConfirmExit)
                StackPanel1.Children.Add(lblSpacer2)
                StackPanel1.Children.Add(cbConfirmRedX)
                StackPanel1.Children.Add(lblSpacer3)
                StackPanel1.Children.Add(cbImportConfirmation)

            Case Is = "Check for update"
                StackPanel1.Children.Add(cbCheckForUpdate)
                StackPanel1.Children.Add(rbDaily)
                StackPanel1.Children.Add(rbWeekly)
                StackPanel1.Children.Add(rbEveryTwoWeeks)
                StackPanel1.Children.Add(rbMonthly)
                StackPanel1.Children.Add(lblSpacer1)
                StackPanel1.Children.Add(btnCheckForUpdateNow)

            Case Is = "Dropbox"
                StackPanel1.Children.Add(cbUseDropbox)
                StackPanel1.Children.Add(lblSpacer1)
                StackPanel1.Children.Add(lblDropboxLine1)
                StackPanel1.Children.Add(tbDropboxPath)
                StackPanel1.Children.Add(lblDropboxLine3)
                StackPanel1.Children.Add(tbDropboxFileName)
                StackPanel1.Children.Add(lblDropboxLine5)
                StackPanel1.Children.Add(tbDropboxDeviceName)

            Case Is = "Imports"
                StackPanel1.Children.Add(cbImportOnByDefault)
                StackPanel1.Children.Add(lblSpacer1)
                StackPanel1.Children.Add(cbImportTag)

            Case Is = "MQTT"

                StackPanel1.Children.Add(cbUseMQTT)
                StackPanel1.Children.Add(lblSpacer1)

                WorkingPanel1.Orientation = Orientation.Horizontal
                WorkingPanel1.Children.Clear()
                WorkingPanel1.Children.Add(lblMQTTBroker)
                WorkingPanel1.Children.Add(tbMQTTBroker)
                StackPanel1.Children.Add(WorkingPanel1)

                WorkingPanel2.Children.Clear()
                WorkingPanel2.Orientation = Orientation.Horizontal
                WorkingPanel2.Children.Add(lblMQTTPort)
                WorkingPanel2.Children.Add(tbMQTTPort)
                StackPanel1.Children.Add(WorkingPanel2)

                StackPanel1.Children.Add(lblSpacer2)

                WorkingPanel3.Children.Clear()
                WorkingPanel3.Orientation = Orientation.Horizontal
                WorkingPanel3.Children.Add(lblMQTTUser)
                WorkingPanel3.Children.Add(tbMQTTUser)
                StackPanel1.Children.Add(WorkingPanel3)

                WorkingPanel4.Children.Clear()
                WorkingPanel4.Orientation = Orientation.Horizontal
                WorkingPanel4.Children.Add(lblMQTTPassword)
                WorkingPanel4.Children.Add(pbMQTTPassword)
                StackPanel1.Children.Add(WorkingPanel4)

                StackPanel1.Children.Add(lblSpacer3)

                WorkingPanel5.Children.Clear()
                WorkingPanel5.Orientation = Orientation.Horizontal
                WorkingPanel5.Children.Add(lblMQTTFilter)
                WorkingPanel5.Children.Add(tbMQTTFilter)
                StackPanel1.Children.Add(WorkingPanel5)

                StackPanel1.Children.Add(lblSpacer4)

                StackPanel1.Children.Add(cbMQTTListenForPayloadOnly)

            Case Is = "Notifications"
                StackPanel1.Children.Add(cbShowNotifications)
                StackPanel1.Children.Add(lblSpacer1)
                StackPanel1.Children.Add(cbShowNotificationResult)
                StackPanel1.Children.Add(cbShowNotificationSource)
                StackPanel1.Children.Add(lblSpacer2)
                StackPanel1.Children.Add(btnTestNotification)
                StackPanel1.Children.Add(lblSpacer3)
                StackPanel1.Children.Add(lblSpacer4)
                StackPanel1.Children.Add(cbIncludeDisconnectAndReconnect)
                SetLookOfNotificationsOptions()

            Case Is = "Pushbullet"
                StackPanel1.Children.Add(cbUsePushbullet)
                StackPanel1.Children.Add(lblSpacer1)
                StackPanel1.Children.Add(lblPushbulletLine1)
                StackPanel1.Children.Add(pbPushBulletAPI)
                StackPanel1.Children.Add(lblPushbulletLine3)
                StackPanel1.Children.Add(tbPushBulletTitleFilter)

            Case Is = "Pushover"
                StackPanel1.Children.Add(cbUsePushover)
                StackPanel1.Children.Add(lblSpacer1)
                StackPanel1.Children.Add(lblPushoverLine1)
                StackPanel1.Children.Add(tbPushoverUserID)
                StackPanel1.Children.Add(lblPushoverLine3)
                StackPanel1.Children.Add(tbPushoverDeviceName)
                StackPanel1.Children.Add(lblPushoverLine5)
                StackPanel1.Children.Add(btnAuthenticate)


            Case Is = "Separating words"
                StackPanel1.Children.Add(lblSeparatingWordsLine1)
                StackPanel1.Children.Add(tbSeparatingWords)

            Case Is = "Session Log"
                StackPanel1.Children.Add(cbAutoScroll)
                StackPanel1.Children.Add(lblSpacer1)
                StackPanel1.Children.Add(cbWriteLogToDisk)
                StackPanel1.Children.Add(lblSpacer2)
                StackPanel1.Children.Add(lblLogFileLocationLine1)
                StackPanel1.Children.Add(btnLogFileLocationLine2)
                StackPanel1.Children.Add(lblLogFileLocationLine3)
                StackPanel1.Children.Add(lblLogFileLocationLine4)
                SetVisibilityOfSessionLogLocation()

            Case Is = "Settings and database files"
                StackPanel1.Children.Add(lblDatabaseLocationLine4)
                StackPanel1.Children.Add(btnDatabaseLocationLine5)
                StackPanel1.Children.Add(lblDatabaseLocationLine3)
                StackPanel1.Children.Add(lblDatabaseLocationLine1)
                StackPanel1.Children.Add(btnDatabaseLocationLine2)
                StackPanel1.Children.Add(lblDatabaseLocationLine6)
                StackPanel1.Children.Add(lblDatabaseLocationLine7)

            Case Is = "Start-up"
                StackPanel1.Children.Add(cbStartPush2RunAtLogon)
                StackPanel1.Children.Add(lblSpacer1)
                StackPanel1.Children.Add(cbStartPush2RunAtLogonAsAdministrator)
                StackPanel1.Children.Add(lblSpacer2)
                StackPanel1.Children.Add(cbShowPush2RunAtStartup)
                StackPanel1.Children.Add(lblSpacer3)
                StackPanel1.Children.Add(cbSuppressStartupNotice)
                StackPanel1.Children.Add(lblSpacer4)
                StackPanel1.Children.Add(cbRequireAPassword)

                If IsAPasswordRequiredForBoss() Then
                    StackPanel1.Children.Add(lblSpacer4)
                    StackPanel1.Children.Add(btnChangeAPassword)
                End If

            Case Is = "Triggers"

                ' ok I know this code is silly, but when I tried to use a textbox I could not get rid of the boarder around it

                Dim DesiredText As String
                DesiredText = "Push2Run may be triggered by Dropbox, MQTT, Pushbullet, Pushover and/or" & vbCrLf
                DesiredText &= "the command line." & vbCrLf
                DesiredText &= vbCrLf
                DesiredText &= "To configure Push2Run for use with Dropbox, MQTT, Pushbullet and/or" & vbCrLf
                DesiredText &= "Pushover, just update the settings for corresponding options below." & vbCrLf
                DesiredText &= vbCrLf
                DesiredText &= "Triggering Push2Run from the command line requires no additional configuration." & vbCrLf

                For Each line In DesiredText.Split(vbCrLf)
                    Dim dummy As New Label
                    dummy.VerticalContentAlignment = VerticalContentAlignment.Top
                    dummy.Margin = New Thickness(0)
                    dummy.Content = line.Trim
                    StackPanel1.Children.Add(dummy)
                Next

            Case Is = "UAC"
                StackPanel1.Children.Add(cbUACLimit)

        End Select

    End Sub

    Private Sub btnDatabaseLocationLine2_Click(sender As Object, e As RoutedEventArgs) Handles btnDatabaseLocationLine5.Click, btnDatabaseLocationLine2.Click, btnLogFileLocationLine2.Click

        'click the button to go to where the database or settings are stored
        Dim parms As String = String.Format("/select, ""{0}""", sender.content.replace("__", "_"))

        gIgnorAutomationOneTime = True
        RunProgram(Environment.GetEnvironmentVariable("windir") & "\explorer.exe", "", parms, False, ProcessWindowStyle.Normal, "")

    End Sub

    Private Sub cbWriteLogToDisk_Checkchanged(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cbWriteLogToDisk.Checked, cbWriteLogToDisk.Unchecked

        SetVisibilityOfSessionLogLocation()

    End Sub

    Private Sub SetVisibilityOfSessionLogLocation()

        If cbWriteLogToDisk.IsChecked Then
            lblLogFileLocationLine1.Visibility = Visibility.Visible
            btnLogFileLocationLine2.Visibility = Visibility.Visible
        Else
            lblLogFileLocationLine1.Visibility = Visibility.Hidden
            btnLogFileLocationLine2.Visibility = Visibility.Hidden
        End If

    End Sub

    Private Sub cbAlwaysOnTop_CheckChanged(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cbAlwaysOnTop.Checked, cbAlwaysOnTop.Unchecked

        MakeTopMost(Process.GetCurrentProcess.Handle, cbAlwaysOnTop.IsChecked)

    End Sub

    Private Sub cbStartPush2RunAtLogon_CheckChanged(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cbStartPush2RunAtLogon.Checked, cbStartPush2RunAtLogon.Unchecked

        If sender.ischecked Then
            StartupShortCut("Add")
            cbStartPush2RunAtLogonAsAdministrator.IsEnabled = True
            cbStartPush2RunAtLogonAsAdministrator.Foreground = cbStartPush2RunAtLogon.Foreground
        Else
            StartupShortCut("Remove")
            cbStartPush2RunAtLogonAsAdministrator.IsChecked = False
            cbStartPush2RunAtLogonAsAdministrator.IsEnabled = False
            cbStartPush2RunAtLogonAsAdministrator.Foreground = Brushes.Gray
        End If

    End Sub


    Private Sub cbByVal(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cbCheckForUpdate.Checked, cbCheckForUpdate.Unchecked

        SetLookOfCheckForUpdate(0)

    End Sub


    Private Sub rbCheckForUpdateFrequency_changed(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles rbDaily.Checked, rbWeekly.Checked, rbEveryTwoWeeks.Checked, rbMonthly.Checked

        SetLookOfCheckForUpdate(sender.tag)

    End Sub


    Private Sub btnCheckForUpdateNow_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnCheckForUpdateNow.Click

        CheckInternetToSeeIfANewVersionIsAvailable(Me, False)

    End Sub


    Private Sub btnAuthenticate_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnAuthenticate.Click

        If RegexUtilities.IsValidEmail(tbPushoverUserID.Text.Trim) Then
        Else
            Dim Result As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "Invalid Pushover user id" & vbCrLf & vbCrLf & "A Pushover user id is the e-mail address you use to log onto Pushover" & vbCrLf & vbCrLf & "The user id entered is not in a standard e-mail address format", "Push2Run - Warning", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK)
            Exit Sub
        End If

        If RegexUtilities.IsValidPushoverDeviceName(tbPushoverDeviceName.Text.Trim) Then
        Else
            Dim Result As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "Invalid Pushover device name" & vbCrLf & vbCrLf & "A Pushover device name may only contain the letters A to Z (upper or lower case), underscores and dashes", "Push2Run - Warning", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK)
            Exit Sub
        End If

        ' not sure why this was here - removed after v4.6.2
        'If RegexUtilities.IsValidDropboxDeviceName(tbDropboxDeviceName.Text.Trim) Then
        'Else
        '    Dim Result As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "Invalid Dropbox device name" & vbCrLf & vbCrLf & "The Dropbox device name may only contain the letters A to Z (upper or lower case), underscores and dashes", "Push2Run - Warning", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK)
        '    Exit Sub
        'End If

        Me.IsEnabled = False

        Dim WindowEnterExistingPassword As WindowEnterExistingPassword = New WindowEnterExistingPassword

        gEnterPasswordWindowTitle = "Push2Run - Enter Pushover password"
        gCurrentOwner = WindowEnterExistingPassword
        WindowEnterExistingPassword.ShowDialog() 'password returns encrypted in gEnteredPassword 

        SeCursor(CursorState.Wait)

        gCurrentOwner = Application.Current.MainWindow

        ' v2.5.1
        My.Settings.PushoverUserID = Me.tbPushoverUserID.Text.Trim

        SetPushoverIdAndSecret(My.Settings.PushoverUserID, EncryptionClass.Decrypt(gEnteredPassword))
        gEnteredPassword = String.Empty

        If ArePushoverIDAndSecretAvailable() Then

            Dim devicename As String = tbPushoverDeviceName.Text.Trim

            SetPushoverDeviceNameAndID(devicename)

            If ArePushoverDeviceNameAndIDAvailable() Then

                DeletePushoverMessages()   ' Delete Messages for newly created device
                SeCursor(CursorState.Normal)
                Dim Result As MessageBoxResult = MessageBox.Show("Authentication OK", "Push2Run - Info", MessageBoxButton.OK, MessageBoxImage.Information, MessageBoxResult.OK)
                gCriticalPushOverErrorReported = False

            Else
                Beep()
                SeCursor(CursorState.Normal)
                Dim Result As MessageBoxResult = MessageBox.Show("Device name did not register", "Push2Run - Warning", MessageBoxButton.OK, MessageBoxImage.Information, MessageBoxResult.OK)

            End If

        Else

            Beep()
            SeCursor(CursorState.Normal)
            Dim Result As MessageBoxResult = MessageBox.Show("Pushover user id, password and/or 2FA code incorrect", "Push2Run - Warning", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK)

        End If

        WindowEnterExistingPassword = Nothing

        Me.IsEnabled = True
        SeCursor(CursorState.Normal)

    End Sub



    Private Sub AddAPassword_CheckChanged(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cbRequireAPassword.Checked

        If LoadUnderway Then Exit Sub

        SeCursor(CursorState.Wait)

        If IsAPasswordRequiredForBoss() Then

            SetPasswordIsRequireFlag(sender.ischecked)

        Else

            Dim WindowNewPassword As WindowNewPassword = New WindowNewPassword

            gCurrentOwner = WindowNewPassword
            WindowNewPassword.ShowDialog()
            gCurrentOwner = Application.Current.MainWindow

            WindowNewPassword = Nothing

            If gPasswordWasCorrectlyEnteredInPasswordWindow Then

                SetPasswordIsRequireFlag(sender.ischecked)

            Else

                LoadUnderway = True
                sender.IsChecked = False
                LoadUnderway = False
                e.Handled = True

            End If

        End If

        BuildStackPanel("Start-up")

        SeCursor(CursorState.Normal)

    End Sub


    Private Sub RemoveAPassword_CheckChanged(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cbRequireAPassword.Unchecked

        If LoadUnderway Then Exit Sub

        If cbRequireAPassword.IsChecked Then Exit Sub

        SeCursor(CursorState.Wait)

        'v1.6 require the password to remove the password
        Dim WindowPromptForPasswordWindow As WindowPromptForPassword = New WindowPromptForPassword

        gCurrentOwner = WindowPromptForPasswordWindow
        WindowPromptForPasswordWindow.ShowDialog()
        gCurrentOwner = Application.Current.MainWindow

        If gPasswordWasCorrectlyEnteredInPasswordWindow Then

            If SetPassword(String.Empty) Then
            Else
                Dim Result As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "Password update error." & vbCrLf & "Password not set.", "Push2Run - Warning", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK)
                Result = TopMostMessageBox(gCurrentOwner, "There was a problem updating your password; your database may have been corrupted.", "Push2Run - Warning", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK)
                gPasswordWasCorrectlyEnteredInPasswordWindow = False
            End If

            SetPasswordIsRequireFlag(sender.ischecked)

        Else
            LoadUnderway = True
            cbRequireAPassword.IsChecked = True
            SetPasswordIsRequireFlag(True)
            LoadUnderway = False
        End If

        WindowPromptForPasswordWindow = Nothing

        BuildStackPanel("Start-up")

        SeCursor(CursorState.Normal)

    End Sub


    Private Sub btnChangePassword_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnChangeAPassword.Click

        Dim WindowChangePassword As WindowChangePassword = New WindowChangePassword

        gCurrentOwner = WindowChangePassword
        WindowChangePassword.ShowDialog()
        gCurrentOwner = Application.Current.MainWindow

        WindowChangePassword = Nothing

    End Sub

    Private Sub tbMQTTPort_KeyDown(sender As Object, e As KeyEventArgs) Handles tbMQTTPort.KeyDown

        Dim kc As New KeyConverter
        Dim Regex = New Regex("[^0-9]+\t")
        e.Handled = Regex.IsMatch(kc.ConvertToInvariantString(e.Key).Replace("NumPad", ""))

    End Sub


    Private Sub WindowOptions_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

        Call IsAPasswordRequiredForBoss() ' this ensures the password required flag is correctly set

    End Sub

    Private Sub ThisWindow_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged

        KeepHelpOnTop()

    End Sub

    Private Sub WindowAddChange_StateChanged(sender As Object, e As EventArgs) Handles Me.StateChanged

        If sender.windowstate = WindowState.Minimized Then
            Me.WindowState = WindowState.Normal
        End If

    End Sub

    Private Sub BtnHelp_Click(sender As Object, e As RoutedEventArgs) Handles BtnHelp.Click

        Me.Dispatcher.Invoke(New OpenAWebPageDelegate(AddressOf OpenAWebPage), New Object() {gWebPageHelpOptionsWindow})

    End Sub

    Private Sub cbShowNotifications_Click(sender As Object, e As RoutedEventArgs) Handles cbShowNotifications.Click

        SetLookOfNotificationsOptions()

    End Sub

    Private Sub SetLookOfNotificationsOptions()

        cbShowNotificationResult.IsEnabled = cbShowNotifications.IsChecked
        cbShowNotificationSource.IsEnabled = cbShowNotifications.IsChecked
        btnTestNotification.IsEnabled = cbShowNotifications.IsChecked

    End Sub

    Private Sub btnTestNotification_Click(sender As Object, e As RoutedEventArgs) Handles btnTestNotification.Click

        Me.IsEnabled = False
        SeCursor(CursorState.Wait)

        ToastNotificationPrimative("Something you might say", "Action completed successfully", "Test request", cbShowNotificationResult.IsChecked, cbShowNotificationSource.IsChecked, Now)

        System.Threading.Thread.Sleep(6500) ' extra long delay to allow notification to appear and dismiss
        Me.IsEnabled = True

        SeCursor(CursorState.Normal)

    End Sub

End Class

