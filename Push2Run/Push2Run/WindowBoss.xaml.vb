'Copyright Rob Latour 2026

'GUI Related
Imports System.Data
Imports System.Data.SQLite
Imports System.IO
Imports System.Net.NetworkInformation
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Windows.Forms
Imports System.Windows.Threading
Imports System.Xml.Serialization ' Does XML serializing for a class.
Imports MQTTnet
Imports MQTTnet.Client
Imports Push2Run.My

Class WindowBoss
    Inherits Window

#Region "GUI Related "

    Private Const MasterControlSwitchID As Integer = 1

    Private Enum ListViewColumns
        ID = 0
        SortOrder = 1
        Enabled = 2
        Description = 3
        ListenFor = 4
        Open = 5
        StartIn = 6
        Parameters = 7
        Admin = 8
        StartingWindowState = 9
        KeysToSend = 10
    End Enum

    Private ViewDescription As Boolean = True
    Private ViewListenFor As Boolean = True
    Private ViewOpen As Boolean = True
    Private ViewParameters As Boolean = True
    Private ViewStartIn As Boolean = True
    Private ViewAdmin As Boolean = True
    Private ViewStartingWindowState As Boolean = True
    Private ViewKeysToSend As Boolean = True

    Private ViewDisabledCards As Boolean = True
    Private ViewFilters As Boolean = True
    Private FilterIsActive As Boolean = False

    Private OriginalParametersColumn As GridViewColumn

    Private OKToClose As Boolean = False

    Private RunThisCommandOnStartup As String = String.Empty

#End Region

    Const CheckForGoodConnectionEverySecond As Integer = 1000 ' used when the connection is good; re-check it frequently so as to be aware quickly if connection goes down 
    Const CheckForGoodConnectionEvery15Seconds As Integer = 15000 ' used when a connection is lost, when the connection is down check this often to see if it can be reconnected

#Region "Pushbullet Related"

    ' ref https://docs.pushbullet.com/
    ' ref http://websocket4net.codeplex.com/

    Const PushbulletServerName As String = "wss://stream.pushbullet.com/websocket/"

    Const AddressForGettingDevices As String = "https://api.pushbullet.com/v2/devices"
    Const AddressForGettingPushes As String = "https://api.pushbullet.com/v2/pushes"
    Const AddressForSendingPushes As String = "https://api.pushbullet.com/v2/pushes"
    Const AddressForDismissingPushes As String = "https://api.pushbullet.com/v2/pushes/{iden}"

    Const AddressForKeepingPushbulletAccountActive As String = "https://zebra.pushbullet.com/"
    Const AddressForGettingPushbulletUserIden As String = "https://api.pushbullet.com/v2/users/me"

    Private LastTimeAPushbulletPushWasRecieved_Unix As Double = 0
    Private Structure NoteInfo
        Dim Iden As String
        Dim Title As String
        Dim Body As String
        Dim Dismissed As Boolean
    End Structure

    Private PushbulletWebSocket As WebSocket4Net.WebSocket
    Private PushoverWebSocket As WebSocket4Net.WebSocket

    Private DisablePushoverProcessing As Boolean = True

#End Region

#Region "MQTT Related"

    ' Private MQTTClient As MqttClient ' moved to modCommon to allow use of publish function in modCommon
    Private MQTTFactory As New MqttFactory

#End Region
    Enum MessageSource

        InternallyGenerated = 0
        CommandLine = 1
        Dropbox = 2
        MQTT = 3
        Pushbullet = 4
        Pushover = 5
        UserRequest = 6

    End Enum

    Private gEarlyShutdown As Boolean = False
    Private gIgnorWindowStateChanges As Integer = True

    Friend WithEvents Timer1 As New System.Windows.Forms.Timer

    Private Sub DeterimineTheStateOfTheEnvironment()

        DetermineIfWindows10OrAbove()
        DetermineIfUACIsOn()
        DetermineIfProgramIsRunningWithAdministrativePrivileges()
        DetermineIfRunningInSandbox()

    End Sub

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        If My.Settings.ApplicationVersion = "not yet set" Then

            'validate by looking for previous settings 

            My.Settings.Upgrade()  ' starting in v4.6 this works differently, hence the kludge logic below
            My.Settings.Save()

        End If

#If DEBUG Then

        Try
            ' clears the visual studio intermediate window

            Dim dte As Object = Marshal.GetActiveObject("VisualStudio.DTE.15.0")
            dte.Windows.Item("Immediate Window").Activate() 'Activate Immediate Window  
            dte.ExecuteCommand("Edit.SelectAll")
            dte.ExecuteCommand("Edit.ClearAll")
            Marshal.ReleaseComObject(dte)

        Catch ex As Exception
        End Try

#End If

        Try

            gCurrentOwner = Application.Current.MainWindow

            'build window off screen
            Me.Top = -5000
            Me.Left = -5000

            'V4.6 
            gSQLiteFullDatabaseName = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData) & "\Push2Run\Push2Run.db3"
            gSQLiteConnectionString = "Data Source=" & gSQLiteFullDatabaseName & ";"

            UpdateSettingsAsNecissary()

            DeterimineTheStateOfTheEnvironment()

            InitializeFileWatcherVariables() ' moved up higher to ensure it runs v4.2

            SetupToManageCards()

            gMyUniqueID = My.Settings.MyUniqueID

            Dim CommandLine As String = Microsoft.VisualBasic.Command.Trim

            Dim CommandLineArray() As String = Environment.GetCommandLineArgs

            gFirstStartOfTheSession = True


            'v4.6
            'gAdminFlag = False
            gAdminFlag = gIsAdministrator

            gForceTheShowingOfTheMainWindowOnRestart = False

            If gRunningInASandbox Then

                If My.Settings.SandboxMessageHasBeenShown Then
                Else

                    Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner,
    "It appears Push2Run is running in the Windows Sandbox." & vbCrLf & vbCrLf &
    "Of note, when a program runs in the Windows Sandbox it runs with administrator privileges." & vbCrLf & vbCrLf &
    "Although Push2Run can run with administrator privileges, normally it does not.  However, as Push2Run appears to be running in the Windows Sandbox it is running with administrative privileges." & vbCrLf & vbCrLf &
    "Having that said, one of Push2Run's features, the ability to allow you to manually switch Push2Run's privilege level between normal and administrative is not available when running in the Windows Sandbox." & vbCrLf & vbCrLf &
    "Also, when Push2Run is first installed it normally provides two example Push2Run cards: one to open Microsoft's Calculator and another to open Microsoft's Notepad program and do some typing into it.  However, as Microsoft's Calculator is unusable in the Windows Sandbox the sample card for it is omitted when installing in the Windows Sandbox." & vbCrLf & vbCrLf &
    "Please click 'OK' below to continue.",
    gThisProgramName & " - in the Windows Sandbox", MessageBoxButton.OK, MessageBoxImage.Asterisk, System.Windows.MessageBoxOptions.None)

                    My.Settings.SandboxMessageHasBeenShown = True
                    My.Settings.Save()

                End If



                gFirstStartOfTheSession = True
                gForceTheShowingOfTheMainWindowOnRestart = True

            Else

                For Each entry In CommandLineArray

                    If entry.StartsWith("StartNormal") Then
                        gFirstStartOfTheSession = True
                        gAdminFlag = False
                        gForceTheShowingOfTheMainWindowOnRestart = False
                        Exit For

                    ElseIf entry.StartsWith("RestartNormal") Then
                        gFirstStartOfTheSession = False
                        gAdminFlag = False
                        gForceTheShowingOfTheMainWindowOnRestart = True
                        Exit For

                    ElseIf entry.StartsWith("StartAdmin") Then
                        gFirstStartOfTheSession = True
                        gAdminFlag = True
                        gForceTheShowingOfTheMainWindowOnRestart = False
                        Exit For

                    ElseIf entry.StartsWith("RestartAdmin") Then
                        gFirstStartOfTheSession = False
                        gAdminFlag = True
                        gForceTheShowingOfTheMainWindowOnRestart = True
                        Exit For

                    End If

                Next

            End If

            Dim TriggerFile As String = Path.GetTempPath & gPush2RunTriggerFileName
            If File.Exists(TriggerFile) Then


                gFirstStartOfTheSession = False
                gAdminFlag = True
                gForceTheShowingOfTheMainWindowOnRestart = True
                File.Delete(TriggerFile)

            End If


            If gFirstStartOfTheSession Then
                If My.Settings.StartBossAsAdministratorByDefault Then
                    gAdminFlag = True
                End If
            End If

            If (Not gIsWindows10OrAbove) AndAlso (Not gIsUACOn) Then

                ' special case, leave everything alone to prevent looping

            Else

                If gFirstStartOfTheSession Then

                    'program is starting for the fist time in a session 

                    If gAdminFlag Then

                        If gIsAdministrator Then
                            ' its all good
                        Else

                            PerformAction("ElevateAtStartup") 'request higher rights

                            gEarlyShutdown = True

                            SafeApplicationCurrentShutdown()

                            Exit Try

                        End If

                    Else

                        'v4.6
                        'If gIsAdministrator Then
                        If gIsUACOn AndAlso gIsAdministrator Then

                            ' request lower rights

                            PerformAction("ChangeAdminRights")

                            gEarlyShutdown = True

                            SafeApplicationCurrentShutdown()

                            Exit Try

                        Else

                            ' its all good
                        End If

                    End If

                Else

                    'program is restarting, but not the fist time in a session 

                    If gAdminFlag Then

                        If gIsAdministrator Then

                            ' its all good
                            If gForceTheShowingOfTheMainWindowOnRestart Then
                            Else
                                gForceTheShowingOfTheMainWindowOnRestart = My.Settings.ShowPush2RunAtStartup
                            End If

                        Else

                            PerformAction("ChangeAdminRights")
                            gEarlyShutdown = True
                            SafeApplicationCurrentShutdown()
                            Exit Try

                        End If

                    Else

                        If gIsAdministrator Then

                            PerformAction("ChangeAdminRights")
                            gEarlyShutdown = True
                            SafeApplicationCurrentShutdown()
                            Exit Try

                        Else

                            ' its all good

                        End If

                    End If

                End If

            End If

            If gFirstStartOfTheSession Then

                If CommandLine.Length > 0 Then
                    RunThisCommandOnStartup = CommandLine.Trim
                End If

                'Ensure this is run only instance of Push2Run running
                Dim AllPush2RunProcesesCount As Integer = Process.GetProcessesByName(Application.ResourceAssembly.GetName.Name).Length

                If AllPush2RunProcesesCount > 1 Then

                    If CommandLine = String.Empty Then

                        CreateACommandToBePickedUpByFileWatcher(CommandToOpenUpMainWindow)

                    Else

                        CreateACommandToBePickedUpByFileWatcher(CommandLine)

                    End If

                    SafeApplicationCurrentShutdown()

                    Exit Try

                End If

            End If

            SetupSystrayIcon()

        Catch ex As Exception

            Dim a As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
            Dim appVersion As Version = a.GetName().Version
            Dim ver As String
            With appVersion
                ver = "v" & .Major & "." & .Minor & "." & .Build & "." & .Revision
            End With

            'v3.6.2
            If ex.Message = "Configuration system failed to initialize" Then
            Else
                MsgBox(" Opps ( " & ver & " ): " & vbCrLf & ex.ToString)
            End If

        End Try

    End Sub
    Private Sub SafeApplicationCurrentShutdown()

        Try
            Application.Current.Shutdown()
        Catch ex As Exception
        End Try

    End Sub


    Private Sub WindowBoss_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded

        'v4.6

        If gEarlyShutdown Then
            Me.WindowState = System.Windows.WindowState.Minimized
            Exit Sub
        End If

        If gSpecialProcessingForTesting Then
            gWebPageVersionCheck = gWebPageVersionCheckWhenTesting
            gWebPageChangeLog = gWebPageChangeLogWhenTesting
            gAutomaticUpdateWebFileName = gAutomaticUpdateWebFileNameWhenTesting
        End If

        Try

            ContextOfMainWindow = Me

            If gBetaVersionInUse Then
                Log("Push2Run version " & My.Application.Info.Version.ToString & " beta started")
            Else
                Log("Push2Run version " & My.Application.Info.Version.ToString & " started")
            End If

            Log("")

            If gIsWindows10OrAbove Then
            Else
                Log("Windows version pre-dates Windows 10")
                Log("")
            End If

            If gIsUACOn Then
                Log("Windows UAC notify feature is on")
            Else
                Log("Windows UAC notify feature is off")
            End If
            Log("")

            If gIsAdministrator Then
                Log("Administrative privileges are enabled")
            Else

                If gIsUACOn Then
                    Log("Normal privileges are in effect")
                Else
                    Log("Administrative privileges are effectively enabled")
                End If

            End If
            Log("")

            Try

                ' auto update stuff

                Dim a As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()

                gThisProgramsPathAndFileName = a.Location

                Dim appVersion As Version = a.GetName().Version
                With gCurrentlyRunningVersion

                    .Major = a.GetName.Version.Major
                    .Minor = a.GetName.Version.Minor
                    .Build = a.GetName.Version.Build
                    .Revision = a.GetName.Version.Revision

                    gVersionInUse = "v" & .Major & "." & .Minor & "." & .Build & "." & .Revision

                End With

                gWebPageHelp = gWebPageHelp.Replace("vX.X.X.X", gVersionInUse)

                gWebPageHelpChangeWindow = gWebPageHelp & "#change"
                gWebPageHelpOptionsWindow = gWebPageHelp & "#options"

                Me.Systray_MenuHeader.Text = "Push2Run " & gVersionInUse

                Me.Systray_MenuHeader.Text = Me.Systray_MenuHeader.Text
                While Me.Systray_MenuHeader.Text.EndsWith(".0")
                    Me.Systray_MenuHeader.Text = Me.Systray_MenuHeader.Text.Remove(Me.Systray_MenuHeader.Text.Length - 2)
                End While

                If File.Exists(gAutomaticUpdateLocalDownloadedFileName) Then

                    If IsDownloadFileForANewerVersion() Then
                    Else
                        File.Delete(gAutomaticUpdateLocalDownloadedFileName)
                    End If

                End If

            Catch ex As Exception

            End Try

            If Process.GetProcessesByName(Application.ResourceAssembly.GetName.Name).Length = 1 Then
                SetupForFileWatcherMonitoring()
            End If

            SetupForNetworkMonitoring()
            StartupShortCut("Ensure")

            CreateDatabaseAsNecissary()

            MenuSort.IsChecked = My.Settings.SortByDescription

            gMenuSort = My.Settings.SortByDescription 'v2.5.3

            ViewFilters = My.Settings.ViewFilters 'v4.8.1
            ViewDisabledCards = My.Settings.ViewDisabledCards 'v4.8.1

            Dim ThePasswordPromptIsNeededAccordingToTheDatabase As Boolean = IsAPasswordRequiredForBoss()

            Me.Visibility = System.Windows.Visibility.Hidden

            If ThePasswordPromptIsNeededAccordingToTheDatabase Then
                Me.WindowState = System.Windows.WindowState.Normal
            Else
                Me.WindowState = System.Windows.WindowState.Minimized
            End If

            SeCursor(CursorState.Wait)

            Me.ShowInTaskbar = False

            ResetEncryptionAndDecriptionToReadAndWrite(ResetEncryptionDecriptionLevel.Data)

            'ensure main window appears on top for initial launch
            MakeTopMost(SafeNativeMethods.FindWindow(Nothing, Me.Title), True)
            If My.Settings.AlwaysOnTop Then
            Else
                MakeTopMost(SafeNativeMethods.FindWindow(Nothing, Me.Title), False)
            End If

            gMasterStatus = GetMasterDesiredStatusFromDatabase()

            'If the master status last time was on pause, set the pause flag
            Me.Systray_MenuPause.Checked = (gMasterStatus = MonitorStatus.Stopped)

            ConfirmSystrayIcons()

            'Determine if a password is required and if so prompt for it

            If ThePasswordPromptIsNeededAccordingToTheDatabase Then

                GetThePassword()

                If gPasswordWasCorrectlyEnteredInPasswordWindow Then

                    ConfirmSystrayIcons()

                    Me.Systray_MenuPasswordRequired.Visible = False
                    Me.Systray_Separator1.Visible = False
                    Me.Systray_MenuShowBoss.Enabled = True
                    Me.Systray_MenuShowOptions.Enabled = True
                    Me.Systray_MenuShowSessionLog.Enabled = True
                    Me.Systray_MenuShowAboutHelp.Enabled = True

                Else

                    OKToClose = True
                    Me.Close()
                    Exit Sub

                End If

            Else

                ConfirmSystrayIcons()

                Me.Systray_MenuPasswordRequired.Visible = False
                Me.Systray_Separator1.Visible = False
                Me.Systray_MenuShowBoss.Enabled = True
                Me.Systray_MenuShowOptions.Enabled = True
                Me.Systray_MenuShowSessionLog.Enabled = True
                Me.Systray_MenuShowAboutHelp.Enabled = True

                gPasswordWasCorrectlyEnteredInPasswordWindow = True

            End If

            SetPriority(ProcessPriorityClass.High)

            If CheckDatabaseVersionAndUpgradeItAsNecissary() Then
            Else

                OKToClose = True

                Me.Close()
                Exit Sub

            End If

            MakeSureMasterRecordIsCorrect()

            AutoCorrectTable1IfNeeded()

            LoadListViewFromDatabase()

            gBossLoadUnderway = True ' LoadListViewFromDatabase sets gBossLoadUnderway = false at its end, so change it back to true while the original load continues below

            ' "View Keys To Send" is handled differently then "View ListenFor" and "View Watch For".
            ' It is handled differently  because "View Keys To Send" data can contain multiple lines of data
            ' and as such the row height of each entry is automatically adjusted to allow for this.
            ' Although setting the column width to zero makes the column disappear, it does not unfortunately
            ' make the row height revert to the default row height - rather it keeps the row height required to 
            ' display the multiple lines.
            ' According the code below removes or adds back in (as needed) the entire column 
            ' this allows the rest of the row to be displayed at the correct height

            ' according: save the OriginalParametersColumn so that it may be removed and re-added later
            Dim gv As GridView = ListView1.View
            OriginalParametersColumn = gv.Columns(ListViewColumns.Parameters)
            gv = Nothing

            ConfirmMasterSwitchAndSystrayIcon()

            LoadWindowLocationSizeAndColumnWidths()

            SetLookOfMenus()

            If My.Settings.CheckForUpdate Then

                If Today < My.Settings.NextVersionCheckDate Then
                Else
                    CheckInternetToSeeIfANewVersionIsAvailable(Me, True)
                End If

            End If

            TurnMasterSwitchOn(gMasterStatus = MonitorStatus.Running)

            AddHandler SessionLogClosed, AddressOf UncheckSessionLogCheckbox
            AddHandler AboutClosed, AddressOf UncheckAboutCheckbox

            gBossLoadUnderway = False

            WorkerTimer.Start()  ' ********* used to reload database 

            ' setup Separating words array for later use
            UpdateGlobalSeparatingWordsArray()

            '*****************************************************************************************************************************

            If My.Settings.FirstTimeRun Then

                My.Settings.FirstTimeRun = False
                My.Settings.Save()

                If TopMostMessageBox(gCurrentOwner, "Welcome!" & vbCrLf & vbCrLf &
                                   "Would you like to visit the Push2Run setup instructions webpage?",
                                   gThisProgramName, MessageBoxButton.YesNo, MessageBoxImage.Asterisk, System.Windows.MessageBoxOptions.None) = MessageBoxResult.Yes Then

                    Call RunProgramStandard(gWebPageSetup, "", "", False, ProcessWindowStyle.Normal, "")

                End If

            End If

            If My.Settings.UseDropbox OrElse My.Settings.UsePushbullet OrElse My.Settings.UsePushover OrElse My.Settings.UseMQTT OrElse My.Settings.SuppressStartupNotice Then

            Else

                If My.Settings.WeclomeMessageHasBeenShown Then
                Else
                    Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner,
                    "Welcome to Push2Run, a program designed to help you control your Windows computer through automation." & vbCrLf & vbCrLf &
                    "To get started Push2Run needs to know how you would like it to work." & vbCrLf & vbCrLf &
                    "This can involve using Dropbox, Pushbullet, Pushover, MQTT, and/or the command line to trigger the things you want done." & vbCrLf & vbCrLf &
                    "Other services, such as those available through Google Assistants products, IFTTT, and Home Assistant can also be used to further automate control." & vbCrLf & vbCrLf &
                    "Telling Push2Run how you would like it to work is done through the program's Options window." & vbCrLf & vbCrLf &
                    "The Options window also includes a 'Help' button.  Clicking it opens Push2Run's help webpage, providing more details on individual settings as well as a link to a first time setup walk through." & vbCrLf & vbCrLf &
                    "I truly hope Push2Run will be of good use to you, and you are welcome to use it for free on as many computers as you would like!" & vbCrLf & vbCrLf &
                    "Rob Latour" & vbCrLf & vbCrLf &
                    "Click 'OK' below to open Push2Run's Options window.",
                    gThisProgramName, MessageBoxButton.OK, MessageBoxImage.Asterisk, System.Windows.MessageBoxOptions.None)

                    My.Settings.WeclomeMessageHasBeenShown = True
                    My.Settings.Save()

                    gOpenOptionsWindowAt = gOpenOptions.AlwaysOnTop
                    OpenOptions()

                End If

            End If

            ' ****************************************************************************************************************************

            GetDrobboxUnderway()

            GetMQTTUnderway()

            GetPushbulletUnderway()

            GetPushoverUnderway()

            '***************************************************************************************************

            SetupStatusBar()

            If My.Settings.ShowPush2RunAtStartup OrElse gForceTheShowingOfTheMainWindowOnRestart Then

                OpenMainWindow()

            End If

            'safe guard against screen size being too big if monitors are switched
            If Me.Width > Forms.Screen.PrimaryScreen.Bounds.Width Then Me.Width = Forms.Screen.PrimaryScreen.Bounds.Width / 2
            If Me.Height > Forms.Screen.PrimaryScreen.Bounds.Height Then Me.Height = Forms.Screen.PrimaryScreen.Bounds.Height / 2


            If gRunningInASandbox Then

                If gIsAdministrator Then
                    If (gIsWindows10OrAbove) AndAlso (Not gIsUACOn) Then
                        Me.Title = "Push2Run (Administrator Privileges) (UAC recommended notify is off) (Sandbox)"
                    Else
                        Me.Title = "Push2Run (Administrator Privileges) (Sandbox)"
                    End If
                Else
                    Me.Title = "Push2Run (Sandbox)"
                End If

                Me.MenuElevate.Header = "Cannot change administrator privileges as program is running in a Sandbox"
                Me.Seperator6a.Visibility = Visibility.Collapsed
                Me.MenuElevate.Visibility = Visibility.Collapsed

            Else

                If (Not gIsWindows10OrAbove) AndAlso (Not gIsUACOn) Then '
                    Me.Title = "Push2Run (UAC recommended notify is off)"
                    Me.MenuElevate.Header = "_Give Push2Run administrator privileges"
                    Me.Seperator6a.Visibility = Visibility.Collapsed
                    Me.MenuElevate.Visibility = Visibility.Collapsed

                ElseIf gIsAdministrator Then
                    Me.Title = "Push2Run (Administrator Privileges)"
                    Me.MenuElevate.Header = "_Remove administrator privileges from Push2Run"

                ElseIf gIsUACOn Then
                    Me.Title = gThisProgramName
                    Me.MenuElevate.Header = "_Give Push2Run administrator privileges"

                Else
                    Me.Title = "Push2Run (UAC recommended notify is off)"
                    Me.MenuElevate.Header = "_Give Push2Run administrator privileges"

                End If

            End If


            If gRunningInASandbox Then
                If Me.Title.Contains(" (Sandbox)") Then
                Else
                    Me.Title &= " (Sandbox)"
                End If
            End If

            If gBetaVersionInUse Then
                Me.Title &= " ( " & gVersionInUse & " beta )"
            End If

            If RunThisCommandOnStartup.StartsWith("StartNormal") OrElse
               RunThisCommandOnStartup.StartsWith("StartAdmin") OrElse
               RunThisCommandOnStartup.StartsWith("RestartNormal") OrElse
               RunThisCommandOnStartup.StartsWith("RestartAdmin") Then

            Else

                If RunThisCommandOnStartup.Length > 0 Then

                    Log("Start-up command line request ...")
                    Log(RunThisCommandOnStartup)

                    Dim FileName As String = RunThisCommandOnStartup.Replace("""", "").Trim


                    If FileName.ToUpper.EndsWith(gPush2RunExtention.ToUpper) Then

                        Dim LoadedCard As CardClass = LoadACard(FileName)

                        With gCurrentlySelectedRow

                            .SortOrder = 15

                            .Description = LoadedCard.Description
                            .ListenFor = LoadedCard.ListenFor
                            .Open = LoadedCard.Open
                            .StartIn = LoadedCard.StartDirectory
                            .Parameters = LoadedCard.Parameters
                            .Admin = LoadedCard.StartWithAdminPrivileges
                            .StartingWindowState = LoadedCard.StartingWindowState
                            .KeysToSend = LoadedCard.KeysToSend

                        End With

                        SafelyAddgCurrentlySelectedRecordIntoTheDatabase(False, FileName)

                        Log("")

                    Else

                        ActionIncomingMessage(MessageSource.CommandLine, RunThisCommandOnStartup)

                    End If

                    RunThisCommandOnStartup = String.Empty

                End If

            End If

            If ViewFilters Then
                tbFilterDescription.Focus()
                DoEvents()
            End If

            'windows automation used to fire an event when a new window is opened and set it to the desired window state
            RegisterForAutomationEvents()

        Catch ex As Exception

            'v3.6.2

            If ex.Message = "Configuration system failed to initialize" Then

                Dim exceptionmessage As String = ex.InnerException.ToString

                Dim Pathname As String = exceptionmessage.Remove(exceptionmessage.IndexOf("user.config"))
                Pathname = Pathname.Remove(0, Pathname.IndexOf(":\") - 1)

                Dim SettingsFilename As String = Pathname & "user.config"
                Dim SettingsFilenameBackup As String = Pathname & "user-1.config"
                Dim ErrorReportFile As String = Pathname & "Error report.txt"

                If System.IO.File.Exists(SettingsFilenameBackup) Then

                    If TopMostMessageBox(gCurrentOwner, "It appears the Push2Run Settings file is corrupt." & vbCrLf & vbCrLf &
                           "Push2Run may be able to recover it." & vbCrLf & vbCrLf &
                           "Would you like Push2Run to try?",
                           gThisProgramName, MessageBoxButton.YesNo, MessageBoxImage.Asterisk, System.Windows.MessageBoxOptions.None) = MessageBoxResult.Yes Then

                        Log("Attempting an automatic recovery")
                        System.IO.File.Delete(SettingsFilename)
                        System.IO.File.Copy(SettingsFilenameBackup, SettingsFilename)
                        Log("Automatic recovery completed")

                        DoEvents()

                        Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "Recovery attempt complete." & vbCrLf & vbCrLf & "Push2Run will automatically restart.", gThisProgramName, MessageBoxButton.OK, MessageBoxImage.Exclamation, System.Windows.MessageBoxOptions.None)

                        Dim sArgument As String
                        Dim sVerb As String

                        If gIsAdministrator Then
                            sArgument = "RestartAdmin"
                            sVerb = "runas"
                        Else
                            sArgument = "Recover"
                            sVerb = "open"
                        End If

                        ReloadPush2Run(sArgument, sVerb)

                        DoEvents()

                    Else

                        My.Computer.FileSystem.WriteAllText(ErrorReportFile, "This file may be safely deleted" & vbCrLf & vbCrLf & Now.ToLongDateString & " " & Now.ToLongTimeString & vbCrLf & vbCrLf & ex.ToString, False)

                        If TopMostMessageBox(gCurrentOwner, "Would you like Push2Run's Settings file folder opened?",
                                    gThisProgramName, MessageBoxButton.YesNo, MessageBoxImage.Asterisk, System.Windows.MessageBoxOptions.None) = MessageBoxResult.Yes Then

                            Dim parms As String = String.Format("/select, ""{0}""", SettingsFilename)
                            Call RunProgramStandard(Environment.GetEnvironmentVariable("windir") & "\explorer.exe", "", parms, False, ProcessWindowStyle.Normal, "")

                        End If

                        System.Threading.Thread.Sleep(1500)

                        SafeNativeMethods.SetForegroundWindow(SafeNativeMethods.FindWindow(Nothing, gThisProgramName)) ' required to make the final pop-up coded (directly below) appear as the topmost screen
                        DoEvents()

                        Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "Push2Run will now close.", "Push2Run - Shutdown", MessageBoxButton.OK, MessageBoxImage.Exclamation, System.Windows.MessageBoxOptions.None)

                        OKToClose = True
                        Me.Close()
                        Exit Sub

                    End If

                End If

            Else

                Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner, ex.ToString, "Push2Run - Fatal Error", MessageBoxButton.OK, MessageBoxImage.Exclamation, System.Windows.MessageBoxOptions.None)

            End If

        End Try


        ' ****************************************************************************************************************************

        Try

            ' ensure CompettingPushThreshold for multiple trigger processing is ok
            If My.Settings.CompetingPushThreshold < 60 Then

                My.Settings.CompetingPushThreshold = 60

            End If

            AddHandler DoAnImport, AddressOf DoImport

            DoBackups()

            TurnBackupTimerOn()

        Catch ex As Exception

        End Try

        UndoDisplayTimer.Interval = 250
        UndoDisplayTimer.Enabled = True

        gIgnorWindowStateChanges = False


        Try

            Dim SettingPathName = System.Configuration.ConfigurationManager.OpenExeConfiguration(System.Configuration.ConfigurationUserLevel.PerUserRoamingAndLocal).FilePath.Replace("user.config", "")
            Dim ErrorReportFile As String = SettingPathName & "Error report.txt"
            If File.Exists(ErrorReportFile) Then
                File.Delete(ErrorReportFile)
            End If

        Catch ex As Exception
        End Try

        SeCursor(CursorState.Normal)

        gInitialStartupUnderway = False

    End Sub


    Private Sub GetDrobboxUnderway()

        Try

            SetupForDropBoxProcessing()

        Catch ex As Exception

            Log("Problem getting Dropbox underway" & vbCrLf & ex.Message.ToString)
            Log("")

        End Try

    End Sub

    Private Async Sub GetMQTTUnderway()

        Try

            If My.Settings.UseMQTT Then

                StartMQTTThreadWaitTillDone()

                Thread.Sleep(100)

                If gMQTTStatusText = "MQTT connected" Then
                    UpdateSubscriptionsWaitUntiDone()
                End If

            End If

        Catch ex As Exception

            Log("Problem getting MQTT underway" & vbCrLf & ex.Message.ToString)
            Log("")

        End Try

    End Sub


    Private Sub StartMQTTThreadWaitTillDone()

        ListOfCurrentlySubscribedTopics.Clear()

        MQTTSetupComplete = False

        StartMQTTThread()

        Dim WaitUntil As DateTime = Now.AddSeconds(5)

        While (Not MQTTSetupComplete) AndAlso (Now < WaitUntil)
            Thread.Sleep(50)
            DoEvents()  'added in v2.0  correct lag if pushbullet access token was not entered
        End While

    End Sub

    Private Sub UpdateSubscriptionsWaitUntiDone()

        UpdateSubscriptionsComplete = False

        UpdateSubscriptions()

        Dim WaitUntil As DateTime = Now.AddSeconds(5)

        While (Not UpdateSubscriptionsComplete) AndAlso (Now < WaitUntil)
            Thread.Sleep(50)
            DoEvents()
        End While

    End Sub

    Private Sub GetPushbulletUnderway()

        Try

            If My.Settings.UsePushbullet Then

                If EncryptionClass.Decrypt(My.Settings.PushBulletAPI) > String.Empty Then

                    If My.Settings.PushBulletTitleFilter > String.Empty Then
                        Call OpenThePushbulletWebSocket()
                    End If

                Else

                    If TopMostMessageBox(gCurrentOwner, "It appears your Pushbullet API is missing." & vbCrLf & vbCrLf & "Would you like to go the options window to check into that?",
                       gThisProgramName, MessageBoxButton.YesNo, MessageBoxImage.Asterisk, System.Windows.MessageBoxOptions.None) = MessageBoxResult.Yes Then
                        gOpenOptionsWindowAt = gOpenOptions.Pushbullet
                        OpenOptions()
                    End If


                End If

                If EncryptionClass.Decrypt(My.Settings.PushBulletAPI) > String.Empty Then
                    If My.Settings.PushBulletTitleFilter > String.Empty Then
                        KeepAccountActiveTimer.Interval = 3000 ' in three seconds ask pushbullet to keep account active
                        KeepAccountActiveTimer.Enabled = True
                        KeepAccountActiveTimer.Start()
                    End If
                End If

            End If

        Catch ex As Exception

            Log("Problem getting Pushbullet underway" & vbCrLf & ex.Message.ToString)
            Log("")

        End Try

        '*****************************************************************************************************************************

        ' Not yet used, keeping code here for potential future use
        ' Dim Devices As String = GetDevices(EncryptDecryptClass.Decrypt(My.Settings.PushBulletAPI))
        ' SendALink(EncryptDecryptClass.Decrypt(My.Settings.PushBulletAPI), "CallClerk", "Here Is a link To the CallClerk website", "http//www.callclerk.com")
        ' SendANote(EncryptDecryptClass.Decrypt(My.Settings.PushBulletAPI), "this Is an example note title", "now Is the time For all good men To come To the aid Of the party")


        ' ****************************************************************************************************************************

    End Sub

    Private Sub GetPushoverUnderway()

        Try

            If My.Settings.UsePushover Then

                If ArePushoverIDAndSecretAvailable() Then

                    If ArePushoverDeviceNameAndIDAvailable() Then

                        'disable processing until the prior Pushover messages have been being deleted
                        DisablePushoverProcessing = True
                        OpenThePushoverWebSocket()
                        DeletePushoverMessages()
                        DisablePushoverProcessing = False

                    Else

                        If TopMostMessageBox(gCurrentOwner, "There appears to be a problem with your Pushover device id." & vbCrLf & vbCrLf & "Would you like to go the options window to check into that?",
                       gThisProgramName, MessageBoxButton.YesNo, MessageBoxImage.Asterisk, System.Windows.MessageBoxOptions.None) = MessageBoxResult.Yes Then
                            gOpenOptionsWindowAt = gOpenOptions.Pushover
                            OpenOptions()
                        End If

                    End If

                Else

                    If TopMostMessageBox(gCurrentOwner, "There appears to be a problem with your Pushover user id or password." & vbCrLf & vbCrLf & "Would you Like to go the options window to check into that?",
                   gThisProgramName, MessageBoxButton.YesNo, MessageBoxImage.Asterisk, System.Windows.MessageBoxOptions.None) = MessageBoxResult.Yes Then
                        gOpenOptionsWindowAt = gOpenOptions.Pushover
                        OpenOptions()
                    End If

                End If

                AddHandler CloseThePushoverWebSocketNow, AddressOf CloseThePushoverWebSocket

            End If

        Catch ex As Exception

            Log("Problem getting Pushover underway" & vbCrLf & ex.Message.ToString)
            Log("")

        End Try

    End Sub

    Private Sub TurnBackupTimerOn()

        Timer1.Interval = MillisecondsToMidnight() + 60000 ' reset time to run tick next at 12:01:00 tomorrow 
        Timer1.Enabled = True
        Timer1.Start()

    End Sub

    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        DoBackups()
        CheckInternetToSeeIfANewVersionIsAvailable(Me, True)

        Timer1.Interval = MillisecondsToMidnight() + 60000 ' reset time to run tick next at 12:01:00 tomorrow 

    End Sub

    Private Sub DoBackups()

        'Database 

        Try

            Dim ReadmeFilename = Path.GetDirectoryName(gSQLiteFullDatabaseName) & "\" & "Read me.txt"

            If File.Exists(ReadmeFilename) Then

                Dim contents As String = File.ReadAllText(ReadmeFilename)

                If contents <> My.Resources.ReadMeDatabaseRestore Then
                    File.Delete(ReadmeFilename)
                    File.WriteAllText(ReadmeFilename, My.Resources.ReadMeDatabaseRestore)
                End If

            Else

                File.WriteAllText(ReadmeFilename, My.Resources.ReadMeDatabaseRestore)

            End If

        Catch ex As Exception
        End Try

        Try

            If File.Exists(gSQLiteFullDatabaseName) Then

                Dim B0_Date, B1_Date As DateTime
                Dim B0, B1, B2, B3, B4, B5, B6, B7 As String

                B0 = gSQLiteFullDatabaseName
                B1 = gSQLiteFullDatabaseName.Replace("Push2Run.db3", "Push2Run-1.db3")

                If File.Exists(B1) Then

                    B0_Date = System.IO.File.GetLastWriteTime(B0)
                    B1_Date = System.IO.File.GetLastWriteTime(B1)

                    If Now.Date = B1_Date.Date Then
                        Log("Daily database backup already run")
                        Exit Try
                    End If

                    If B0_Date = B1_Date Then
                        Log("Daily database backup not run as the database remains unchanged since its last backup")
                        Exit Try
                    End If

                End If

                B2 = gSQLiteFullDatabaseName.Replace("Push2Run.db3", "Push2Run-2.db3")
                B3 = gSQLiteFullDatabaseName.Replace("Push2Run.db3", "Push2Run-3.db3")
                B4 = gSQLiteFullDatabaseName.Replace("Push2Run.db3", "Push2Run-4.db3")
                B5 = gSQLiteFullDatabaseName.Replace("Push2Run.db3", "Push2Run-5.db3")
                B6 = gSQLiteFullDatabaseName.Replace("Push2Run.db3", "Push2Run-6.db3")
                B7 = gSQLiteFullDatabaseName.Replace("Push2Run.db3", "Push2Run-7.db3")

                If File.Exists(B7) Then File.Delete(B7)
                If File.Exists(B6) Then Microsoft.VisualBasic.FileIO.FileSystem.RenameFile(B6, Path.GetFileName(B7))
                If File.Exists(B5) Then Microsoft.VisualBasic.FileIO.FileSystem.RenameFile(B5, Path.GetFileName(B6))
                If File.Exists(B4) Then Microsoft.VisualBasic.FileIO.FileSystem.RenameFile(B4, Path.GetFileName(B5))
                If File.Exists(B3) Then Microsoft.VisualBasic.FileIO.FileSystem.RenameFile(B3, Path.GetFileName(B4))
                If File.Exists(B2) Then Microsoft.VisualBasic.FileIO.FileSystem.RenameFile(B2, Path.GetFileName(B3))
                If File.Exists(B1) Then Microsoft.VisualBasic.FileIO.FileSystem.RenameFile(B1, Path.GetFileName(B2))
                File.Copy(B0, B1)

                Log("Daily database backup completed")

            End If

        Catch ex As Exception

            Log("Daily database backup failed")

        End Try

        Log("")


        ' Settings file

        Try

            Dim ReadmeFilename As String = Path.GetDirectoryName(System.Configuration.ConfigurationManager.OpenExeConfiguration(System.Configuration.ConfigurationUserLevel.PerUserRoamingAndLocal).FilePath) & "\" & "Read me.txt"

            If File.Exists(ReadmeFilename) Then

                Dim contents As String = File.ReadAllText(ReadmeFilename)

                If contents <> My.Resources.ReadMeSettingRestore Then
                    File.Delete(ReadmeFilename)
                    File.WriteAllText(ReadmeFilename, My.Resources.ReadMeSettingRestore)
                End If

            Else

                File.WriteAllText(ReadmeFilename, My.Resources.ReadMeSettingRestore)

            End If

        Catch ex As Exception
        End Try

        Try

            'backup settings

            Dim SettingsFile As String = System.Configuration.ConfigurationManager.OpenExeConfiguration(System.Configuration.ConfigurationUserLevel.PerUserRoamingAndLocal).FilePath

            If File.Exists(SettingsFile) Then

                Dim B0_Date, B1_Date As DateTime
                Dim B0, B1, B2, B3, B4, B5, B6, B7 As String

                B0 = SettingsFile
                B1 = SettingsFile.Replace("user.config", "user-1.config")

                If File.Exists(B1) Then

                    B0_Date = System.IO.File.GetLastWriteTime(B0)
                    B1_Date = System.IO.File.GetLastWriteTime(B1)

                    If Now.Date = B1_Date.Date Then
                        Log("Daily settings backup already run")
                        Exit Try
                    End If

                    If B0_Date = B1_Date Then
                        Log("Daily settings backup not run as the settings file remains unchanged since its last backup")
                        Exit Try
                    End If

                End If

                B2 = SettingsFile.Replace("user.config", "user-2.config")
                B3 = SettingsFile.Replace("user.config", "user-3.config")
                B4 = SettingsFile.Replace("user.config", "user-4.config")
                B5 = SettingsFile.Replace("user.config", "user-5.config")
                B6 = SettingsFile.Replace("user.config", "user-6.config")
                B7 = SettingsFile.Replace("user.config", "user-7.config")

                If File.Exists(B7) Then File.Delete(B7)
                If File.Exists(B6) Then Microsoft.VisualBasic.FileIO.FileSystem.RenameFile(B6, Path.GetFileName(B7))
                If File.Exists(B5) Then Microsoft.VisualBasic.FileIO.FileSystem.RenameFile(B5, Path.GetFileName(B6))
                If File.Exists(B4) Then Microsoft.VisualBasic.FileIO.FileSystem.RenameFile(B4, Path.GetFileName(B5))
                If File.Exists(B3) Then Microsoft.VisualBasic.FileIO.FileSystem.RenameFile(B3, Path.GetFileName(B4))
                If File.Exists(B2) Then Microsoft.VisualBasic.FileIO.FileSystem.RenameFile(B2, Path.GetFileName(B3))
                If File.Exists(B1) Then Microsoft.VisualBasic.FileIO.FileSystem.RenameFile(B1, Path.GetFileName(B2))
                File.Copy(B0, B1)

                Log("Daily settings backup completed")

            End If

        Catch ex As Exception

            Log("Daily settings backup failed")

        End Try

        Log("")

    End Sub


    Private Function MillisecondsToMidnight() As Integer

        Dim ReturnValue As Integer
        Dim ts As TimeSpan

        Dim Tomorrow As DateTime = Today.AddDays(1)
        ts = Tomorrow.Subtract(Now)

        ReturnValue = ts.TotalMilliseconds()

        Return ReturnValue

    End Function


#Region "Systray Icon"

    Public WithEvents SysTrayIcon As Forms.NotifyIcon = New Forms.NotifyIcon
    Public WithEvents Systray_ContextMenu As Forms.ContextMenuStrip = New Forms.ContextMenuStrip

    Public WithEvents Systray_MenuHeader As Forms.ToolStripMenuItem = New Forms.ToolStripMenuItem
    Public WithEvents Systray_Separator1 As Forms.ToolStripSeparator = New Forms.ToolStripSeparator

    Public WithEvents Systray_MenuShowBoss As Forms.ToolStripMenuItem = New Forms.ToolStripMenuItem
    Public WithEvents Systray_MenuShowOptions As Forms.ToolStripMenuItem = New Forms.ToolStripMenuItem
    Public WithEvents Systray_MenuShowSessionLog As Forms.ToolStripMenuItem = New Forms.ToolStripMenuItem
    Public WithEvents Systray_MenuShowAboutHelp As Forms.ToolStripMenuItem = New Forms.ToolStripMenuItem
    Public WithEvents Systray_Separator2 As Forms.ToolStripSeparator = New Forms.ToolStripSeparator

    Public WithEvents Systray_MenuPause As Forms.ToolStripMenuItem = New Forms.ToolStripMenuItem
    Public WithEvents Systray_MenuPasswordRequired As Forms.ToolStripMenuItem = New Forms.ToolStripMenuItem
    Public WithEvents Systray_Separator3 As Forms.ToolStripSeparator = New Forms.ToolStripSeparator
    Public WithEvents Systray_MenuExit As Forms.ToolStripMenuItem = New Forms.ToolStripMenuItem


    Private Sub SetupSystrayIcon()

        SysTrayIcon.ContextMenuStrip = Systray_ContextMenu
        SysTrayIcon.Text = gThisProgramName
        SysTrayIcon.Visible = True

        Me.Systray_ContextMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() _
        {Systray_MenuHeader,
        Systray_Separator1,
        Systray_MenuShowBoss,
        Systray_MenuShowOptions,
        Systray_MenuShowSessionLog,
        Systray_MenuShowAboutHelp,
        Systray_Separator2,
        Systray_MenuPasswordRequired,
        Systray_MenuPause,
        Systray_Separator3,
        Systray_MenuExit})

        Me.Systray_ContextMenu.Name = "MenuContext"
        Me.Systray_ContextMenu.Size = New System.Drawing.Size(175, 104)

        Me.Systray_MenuHeader.Name = "MenuHeader"
        Me.Systray_MenuHeader.Size = New System.Drawing.Size(174, 22)
        Me.Systray_MenuHeader.Text = "Push2Run v(loaded in code)"
        Me.Systray_MenuHeader.Enabled = False

        Me.Systray_Separator1.Name = "ToolStripSeparator1"
        Me.Systray_Separator1.Size = New System.Drawing.Size(171, 6)

        Me.Systray_MenuShowBoss.Name = "MenuShowBoss"
        Me.Systray_MenuShowBoss.Size = New System.Drawing.Size(174, 22)
        Me.Systray_MenuShowBoss.Text = "Main Window"

        Me.Systray_MenuShowOptions.Name = "MenuShowOptions"
        Me.Systray_MenuShowOptions.Size = New System.Drawing.Size(174, 22)
        Me.Systray_MenuShowOptions.Text = "Options"

        Me.Systray_MenuShowSessionLog.Name = "MenuShowSessionLog"
        Me.Systray_MenuShowSessionLog.Size = New System.Drawing.Size(174, 22)
        Me.Systray_MenuShowSessionLog.Text = "Session Log"

        Me.Systray_MenuShowAboutHelp.Name = "MenuShowAboutHelp"
        Me.Systray_MenuShowAboutHelp.Size = New System.Drawing.Size(174, 22)
        Me.Systray_MenuShowAboutHelp.Text = "About/Help"

        Me.Systray_Separator2.Name = "ToolStripSeparator2"
        Me.Systray_Separator2.Size = New System.Drawing.Size(171, 6)

        Me.Systray_MenuPasswordRequired.Name = "MenuPasswordRequired"
        Me.Systray_MenuPasswordRequired.Size = New System.Drawing.Size(174, 22)
        Me.Systray_MenuPasswordRequired.Text = "Enter Password"

        Me.Systray_MenuPause.CheckOnClick = True
        Me.Systray_MenuPause.Name = "MenuPause"
        Me.Systray_MenuPause.Size = New System.Drawing.Size(174, 22)
        Me.Systray_MenuPause.Text = "Pause"

        Me.Systray_Separator3.Name = "ToolStripSeparator2"
        Me.Systray_Separator3.Size = New System.Drawing.Size(171, 6)

        Me.Systray_MenuExit.Name = "MenuExit"
        Me.Systray_MenuExit.Size = New System.Drawing.Size(174, 22)
        Me.Systray_MenuExit.Text = "Exit"

    End Sub

#End Region

#Region "Monitor Network Status"

    Private Shared AdapterToMonitor As String = String.Empty

    Private Sub SetupForNetworkMonitoring()

        AddHandler NetworkChange.NetworkAddressChanged, AddressOf AddressChangedCallback
        AddHandler NetworkChange.NetworkAvailabilityChanged, AddressOf AvailabilityChangedCallback

        gNetworkMonitoringOn = True

        ReportCurrentNetworkStatus()

    End Sub

    Private Sub AddressChangedCallback(ByVal sender As Object, ByVal e As EventArgs)

        ReportCurrentNetworkStatus()

        If CurrentNetworkStatus <> OperationalStatus.Up Then
            ResetTheTimePushbulletWasLastAccessed(Now.AddDays(-1))
        End If

    End Sub

    Private Sub AvailabilityChangedCallback(ByVal sender As Object, ByVal e As EventArgs)

        ReportCurrentNetworkStatus()

        If CurrentNetworkStatus <> OperationalStatus.Up Then
            ResetTheTimePushbulletWasLastAccessed(Now.AddDays(-1))
        End If

    End Sub

    Private CurrentNetworkStatus As OperationalStatus = OperationalStatus.Unknown
    Private Sub ReportCurrentNetworkStatus()

        Static LastKnownStatus As OperationalStatus = OperationalStatus.Testing

        Dim DeviceFound As Boolean = False

        Dim CurrentAdapterToMonitor As String = FindDefaultNetworkAdapterName() 'v2.1.2

        If AdapterToMonitor = CurrentAdapterToMonitor Then
        Else
            AdapterToMonitor = CurrentAdapterToMonitor
            Log("Current network adapter is " & AdapterToMonitor)
            Log("")
        End If

        For Each nic As NetworkInterface In NetworkInterface.GetAllNetworkInterfaces()

            If nic.Name = AdapterToMonitor Then
                CurrentNetworkStatus = nic.OperationalStatus
                DeviceFound = True
                Exit For
            End If

        Next nic

        If DeviceFound Then
        Else
            CurrentNetworkStatus = OperationalStatus.NotPresent
        End If

        If CurrentNetworkStatus <> LastKnownStatus Then

            Dim ReportedNetworkStatus As String = CurrentNetworkStatus.ToString.ToLower.Replace("notpresent", "disabled")

            Log("Network status is " & ReportedNetworkStatus)
            Log("")

            'added if in v3.7.1 to prevent duplicate notifications
            If CurrentNetworkStatus = gPreviousNetworkConnectionStatus Then
            Else
                ToastNotificationForNetorkEvents("Network status", ReportedNetworkStatus, Now)
                gPreviousNetworkConnectionStatus = CurrentNetworkStatus
            End If

            'v1.3 
            If (CurrentNetworkStatus = OperationalStatus.Up) AndAlso WebSocketErrorWasReportedPushbullet Then
                WebSocketErrorWasReportedPushbullet = False
                Call OpenThePushbulletWebSocket()
            End If

            LastKnownStatus = CurrentNetworkStatus

        End If

    End Sub
    Private Function FindDefaultNetworkAdapterName() As String

        Dim ReturnValue As String = String.Empty

        Dim MyAddressList As New List(Of String)

        'replaced the commented code below with the revise code underneath it using GetHostByName was throwing a compiler warning about being obsolete 'v3.7.1
        'For Each address As System.Net.IPAddress In System.Net.Dns.GetHostByName(System.Net.Dns.GetHostName()).AddressList
        '    MyAddressList.Add(address.ToString)
        'Next

        For Each address As System.Net.IPAddress In System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName()).AddressList
            MyAddressList.Add(address.ToString)
        Next

        Dim adapters As NetworkInterface() = NetworkInterface.GetAllNetworkInterfaces()
        Dim nic As NetworkInterface

        For Each nic In adapters

            Dim ipProps As IPInterfaceProperties = nic.GetIPProperties()

            For Each obj As UnicastIPAddressInformation In ipProps.UnicastAddresses

                If MyAddressList.Contains(obj.Address.ToString) Then
                    ReturnValue = nic.Name
                    Exit For
                End If

            Next

            If ReturnValue > String.Empty Then Exit For

        Next nic

        Return ReturnValue

    End Function

#End Region

    Private Function OpenThePushbulletWebSocket() As Boolean

        If PushbulletEnbledInTesting Then
        Else
            Return False
        End If

        If My.Settings.UsePushbullet Then
        Else
            Return False
        End If

        Dim ReturnValue As Boolean = False

        Try

            CloseThePushbulletWebSocket()

            PushbulletWebSocket = New WebSocket4Net.WebSocket(PushbulletServerName & EncryptionClass.Decrypt(My.Settings.PushBulletAPI), sslProtocols:=System.Security.Authentication.SslProtocols.Tls12) 'v2.1.2

            AddHandler PushbulletWebSocket.Opened, AddressOf Websocket_Opened_Pushbullet
            AddHandler PushbulletWebSocket.MessageReceived, AddressOf Websocket_MessageReceived_Pushbullet
            AddHandler PushbulletWebSocket.DataReceived, AddressOf Websocket_DataReceived_Pushbullet
            AddHandler PushbulletWebSocket.Error, AddressOf Websocket_Error_Pushbullet
            AddHandler PushbulletWebSocket.Closed, AddressOf Websocket_Closed_Pushbullet

            PushbulletWebSocket.Open()

            Dim TenSecondsFromNow As Date = Now.AddSeconds(10)
            While (PushbulletWebSocket.State = WebSocket4Net.WebSocketState.Connecting) AndAlso (Now < TenSecondsFromNow)
                Thread.Sleep(100)
                DoEvents()  'added in v2.0  correct lag if pushbullet access token was not entered
            End While

            If PushbulletWebSocket.State = PushbulletWebSocket.State.Open Then

                LastTimeAPushbulletPushWasRecieved_Unix = GetServerTimeOfMostRecentPush_Unix()
                LastTimeDataWasReceivedFromPushbullet = Now
                ReturnValue = True
                Log("Pushbullet connected")
                Log("")

                'added if in v3.7.1 to prevent duplicate notifications
                If gPushbulletConnectionStatus = ConnectionStatus.Connected Then
                Else
                    gPushbulletConnectionStatus = ConnectionStatus.Connected
                    ToastNotificationForNetorkEvents("Pushbullet", "Connection established", Now)
                End If

                ResetTheTimePushbulletWasLastAccessed(Now)

            Else

                LastTimeAPushbulletPushWasRecieved_Unix = GetUnixTime(Now.AddDays(-1))
                LastTimeDataWasReceivedFromPushbullet = Now.AddDays(-1)
                ReturnValue = False
                Log("Pushbullet not connected")
                Log("")

                If gPushbulletConnectionStatus = ConnectionStatus.DisconnectedorClosed Then
                Else
                    gPushbulletConnectionStatus = ConnectionStatus.DisconnectedorClosed
                    ToastNotificationForNetorkEvents("Pushbullet", "Connection not established", Now)
                End If

                ResetTheTimePushbulletWasLastAccessed(Now.AddDays(-1))

            End If

        Catch ex As Exception
            Log(ex.Message.ToString)
            Log("")
        End Try

        Return ReturnValue

    End Function

    Private Sub CloseThePushbulletWebSocket()

        Try

            If PushbulletWebSocket IsNot Nothing Then

                If PushbulletWebSocket.State = PushbulletWebSocket.State.Open Then

                    PushbulletWebSocket.Close()

                    'turns out you have to remove the old handlers before recreating new ones

                    RemoveHandler PushbulletWebSocket.Opened, AddressOf Websocket_Opened_Pushbullet
                    RemoveHandler PushbulletWebSocket.MessageReceived, AddressOf Websocket_MessageReceived_Pushbullet
                    RemoveHandler PushbulletWebSocket.DataReceived, AddressOf Websocket_DataReceived_Pushbullet
                    RemoveHandler PushbulletWebSocket.Error, AddressOf Websocket_Error_Pushbullet
                    RemoveHandler PushbulletWebSocket.Closed, AddressOf Websocket_Closed_Pushbullet

                    Log("Pushbullet connection closed")
                    Log("")

                    gPushbulletConnectionStatus = ConnectionStatus.DisconnectedorClosed
                    ToastNotificationForNetorkEvents("Pushbullet", "Connection closed", Now)

                End If

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub OpenThePushoverWebSocket()

        If gCriticalPushOverErrorReported Then Exit Sub

        Try

            If PushoverEnabledInTesting Then
            Else
                Exit Try
            End If

            If ArePushoverIDAndSecretAvailable() Then
            Else
                Exit Try
            End If

            If ArePushoverDeviceIDAndSecretAvailable() Then
            Else
                Log("Connection with Pushover cannot be established")
                Log("The problem may be with your Pushover ID/password or your Pushover device")
                Log("")

                ToastNotificationForNetorkEvents("Pushover", "Connection cannot be established", Now)

                Exit Try

            End If

            CloseThePushoverWebSocket()

            PushoverWebSocket = New WebSocket4Net.WebSocket("wss://client.pushover.net/push", sslProtocols:=System.Security.Authentication.SslProtocols.Tls12)

            AddHandler PushoverWebSocket.Opened, AddressOf Websocket_Opened_Pushover
            AddHandler PushoverWebSocket.MessageReceived, AddressOf Websocket_MessageReceived_Pushover
            AddHandler PushoverWebSocket.DataReceived, AddressOf Websocket_DataReceived_Pushover
            AddHandler PushoverWebSocket.Error, AddressOf Websocket_Error_Pushover
            AddHandler PushoverWebSocket.Closed, AddressOf Websocket_Closed_Pushover

            PushoverWebSocket.Open()

            Dim TenSecondsFromNow As Date = Now.AddSeconds(10)
            While (PushoverWebSocket.State = WebSocket4Net.WebSocketState.Connecting) AndAlso (Now < TenSecondsFromNow)
                Thread.Sleep(100)
                DoEvents()
            End While

            If PushoverWebSocket.State = PushoverWebSocket.State.Open Then

                LastTimeDataWasReceivedFromPushover = Now

                Dim LogonCredentials As String = "login:" & EncryptionClass.Decrypt(My.Settings.PushoverDeviceID) & ":" & EncryptionClass.Decrypt(My.Settings.PushoverSecret) & vbLf
                PushoverWebSocket.Send(LogonCredentials)

                Log("Pushover connected")
                Log("")

                'added if in v3.7.1 to prevent duplicate notifications
                If Global.Push2Run.modCommon.gPushoverConnectionStatus = Global.Push2Run.modCommon.ConnectionStatus.Connected Then
                Else
                    modCommon.gPushoverConnectionStatus = ConnectionStatus.Connected
                    ToastNotificationForNetorkEvents("Pushover", "Connection established", Now)
                End If

            Else

                LastTimeDataWasReceivedFromPushover = Now.AddDays(-1)
                Log("Pushover not connected")
                Log("")

                'added if in v3.7.1 to prevent duplicate notifications
                If Global.Push2Run.modCommon.gPushoverConnectionStatus = Global.Push2Run.modCommon.ConnectionStatus.DisconnectedorClosed Then
                Else
                    modCommon.gPushoverConnectionStatus = ConnectionStatus.DisconnectedorClosed
                    ToastNotificationForNetorkEvents("Pushover", "Connection not established", Now)
                End If


            End If

        Catch ex As Exception

            Log(ex.Message.ToString)
            Log("")

        End Try

    End Sub

    Private Sub CloseThePushoverWebSocket()

        Try

            If PushoverWebSocket IsNot Nothing Then

                If PushoverWebSocket.State = PushoverWebSocket.State.Open Then

                    PushoverWebSocket.Close()

                    RemoveHandler PushoverWebSocket.Opened, AddressOf Websocket_Opened_Pushover
                    RemoveHandler PushoverWebSocket.MessageReceived, AddressOf Websocket_MessageReceived_Pushover
                    RemoveHandler PushoverWebSocket.DataReceived, AddressOf Websocket_DataReceived_Pushover
                    RemoveHandler PushoverWebSocket.Error, AddressOf Websocket_Error_Pushover
                    RemoveHandler PushoverWebSocket.Closed, AddressOf Websocket_Closed_Pushover

                    LastTimeDataWasReceivedFromPushover = Now.AddDays(-1)

                    Log("Connection with Pushover closed")
                    Log("")

                    modCommon.gPushoverConnectionStatus = ConnectionStatus.DisconnectedorClosed
                    ToastNotificationForNetorkEvents("Pushover", "Connection closed", Now)

                End If

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private ThereAreStillMoreMessagesToProcess As Boolean = True
    Private PushoverStillProssessingFlag As Boolean = True

    Private Sub ProcessANewPushoverMessage()

        If PushoverAutoBan() Then Exit Sub

        PushoverStillProssessingFlag = True

        Try

            Dim Status As String = String.Empty
            Dim ServerResponse As String = String.Empty

            SendRequest("", "GET", "https://api.pushover.net/1/messages.json" & "?secret=" & EncryptionClass.Decrypt(My.Settings.PushoverSecret) & "&device_id=" & EncryptionClass.Decrypt(My.Settings.PushoverDeviceID), String.Empty, ServerResponse)

            ' code assumes only one message will be need to be processed; api can return more than one, but that should not happen

            Status = GetFirstMatchingValueFromJSONResponseString("status", ServerResponse)

            If Status = "1" Then

                Dim Message As String = GetFirstMatchingValueFromJSONResponseString("message", ServerResponse).Trim

                'Action Pushover message
                Log("")
                Log("Incoming Pushover push ...")

                If My.Settings.UsePushover Then
                    If Message = String.Empty Then
                        Log("Message was empty - no further action will be taken")
                        Log("")
                    Else
                        Log(Message)
                        ActionIncomingMessage(MessageSource.Pushover, Message)
                    End If
                Else
                    Log("Pushover not enabled in options - no further action will be taken")
                End If

                'Delete Pushover message so that it is not processed again

                PushoverMessageIdToDelete = GetFirstMatchingValueFromJSONResponseString("id_str", ServerResponse)
                DeletePushoverMessageID_Processing = String.Empty

                Dim NewThread2 As Thread = New Thread(AddressOf DeletePushoverMessagesID_Async)
                NewThread2.Start()

                While DeletePushoverMessageID_Processing = String.Empty
                    Thread.Sleep(100)
                End While

            Else

                Log("Pushover reported a problem with the use of your device! (a)")
                Log("")

                gCriticalPushOverErrorReported = True
                LastTimeDataWasReceivedFromPushover = Now.AddDays(-1)
                CloseThePushoverWebSocket()
                ThereAreStillMoreMessagesToProcess = False

            End If

        Catch ex As Exception

            Log("Pushover processing encounter an unexpected error!")
            Log(ex.Message.ToString)
            Log("")

            gCriticalPushOverErrorReported = True
            LastTimeDataWasReceivedFromPushover = Now.AddDays(-1)
            CloseThePushoverWebSocket()
            ThereAreStillMoreMessagesToProcess = False

            MsgBox(ex)

        End Try

        PushoverStillProssessingFlag = False

    End Sub

    Private Function TruncateTime(ByVal dateTime As DateTime, ByVal timeSpan As TimeSpan) As DateTime
        If timeSpan = TimeSpan.Zero Then Return dateTime
        If dateTime = DateTime.MinValue OrElse dateTime = DateTime.MaxValue Then Return dateTime
        Return dateTime.AddTicks(-(dateTime.Ticks Mod timeSpan.Ticks))
    End Function

    Friend Structure PushoverAutoBanRecord
        Dim timeStamp As DateTime
        Dim count As Integer
    End Structure

    '                                                                  Threshold is more than
    Const PushoverAutoBanThresholdTansactionLimit As Integer = 90     ' 90 requests in 
    Const PushoverAutoBanTimeFrameThreshold As Integer = 3600         ' one hour  (if changed from 1 hour, log displays below will need to change too)

    ' there are two transactions in each request
    Private Function PushoverAutoBan() As Boolean

        'returns True when the total number of times this routine has been called within the PushoverAutoBanTimeFrameThreshold exceeds the PushoverAutoBanThresholdLimit
        'once the routine returns true, it will continue to return true until the program is restarted

        Dim ReturnValue As Boolean = False

        Static AutoBanHasBeenSet As Boolean = False

        Static AutoBanTable(PushoverAutoBanTimeFrameThreshold) As PushoverAutoBanRecord

        Try

            If AutoBanHasBeenSet Then

                'once the autoban has been set the only way to reset it is to exit the program and start it up again
                ReturnValue = True

            Else

                Dim currentTime As DateTime = TruncateTime(Now, TimeSpan.FromSeconds(1)) ' truncate off milliseconds

                Dim timeFrameToMonitor = Now.AddSeconds(-1 * PushoverAutoBanTimeFrameThreshold)

                ' account for the number of times this routine has been called by adding 1 within in the AutoBanTable 

                ' Given it will be possible to find a entry in the AutoBanTable where the current time either
                '    a) matches a timestamp in the table, or 
                '    b) can be added into the table replacing a timestamp older than the timeframe being monitored
                '
                ' then 
                '
                '   a) 1 to the matching entries count
                '
                '   b) set timestamp of the entry to be replace with the current timestamp and set its count to 1

                Dim MatchFound As Boolean = False

                For x As Integer = 0 To PushoverAutoBanTimeFrameThreshold - 1
                    If AutoBanTable(x).timeStamp = currentTime Then
                        AutoBanTable(x).count += 1
                        MatchFound = True
                        Exit For
                    End If
                Next

                If Not MatchFound Then
                    For x As Integer = 0 To PushoverAutoBanTimeFrameThreshold - 1
                        If AutoBanTable(x).timeStamp < timeFrameToMonitor Then
                            AutoBanTable(x).timeStamp = currentTime
                            AutoBanTable(x).count = 1
                            Exit For
                        End If
                    Next
                End If

                ' calculate a total of the counts associated with all entries in the AutoBanTable within the time frame being monitored

                Dim totalHitsInTheTimeFrame = 0

                For x As Integer = 0 To PushoverAutoBanTimeFrameThreshold - 1
                    If AutoBanTable(x).timeStamp >= timeFrameToMonitor Then
                        totalHitsInTheTimeFrame += AutoBanTable(x).count
                    End If
                Next

                If totalHitsInTheTimeFrame > PushoverAutoBanThresholdTansactionLimit Then

                    ReturnValue = True
                    AutoBanHasBeenSet = True
                    Log("!!! Pushover limit of " & PushoverAutoBanThresholdTansactionLimit.ToString & " request per hour exceeded !!!")
                    Log("!!! If you have not issued over " & PushoverAutoBanThresholdTansactionLimit.ToString & " request in the last hour please contact info@push2run !!!")
                    Log("")

                    gCriticalPushOverErrorReported = True
                    LastTimeDataWasReceivedFromPushover = Now.AddDays(-1)
                    CloseThePushoverWebSocket()

                    'Else

                    '    Log("Pushover hits in the threshold period = " & totalHitsInTheTimeFrame)

                End If

            End If

        Catch ex As Exception

            ReturnValue = True
            AutoBanHasBeenSet = True

        End Try

        Return ReturnValue

    End Function

#Region "GUI Related"

    'gFirstStartOfTheSession = true for 1st start of the session; false for restarting (all subsequent starts)

    Private gFirstStartOfTheSession As Boolean = True
    Private gForceTheShowingOfTheMainWindowOnRestart As Boolean = False
    Private gAdminFlag As Boolean = False


    Friend Sub GetThePassword()

        Try

            Me.Systray_MenuPasswordRequired.Enabled = False

            'have the user enter the password now
            Dim WindowPromptForPasswordWindow As WindowPromptForPassword = New WindowPromptForPassword

            gCurrentOwner = WindowPromptForPasswordWindow
            WindowPromptForPasswordWindow.ShowDialog()
            gCurrentOwner = Application.Current.MainWindow

            If gPasswordWasCorrectlyEnteredInPasswordWindow Then

                Me.Systray_MenuPasswordRequired.Visible = False
                Me.Systray_Separator1.Visible = False
                Me.Systray_MenuShowBoss.Enabled = True
                Me.Systray_MenuShowOptions.Enabled = True
                Me.Systray_MenuShowSessionLog.Enabled = True
                Me.Systray_MenuShowAboutHelp.Enabled = True

            Else

                Me.Systray_MenuPasswordRequired.Enabled = True

            End If

            ConfirmMasterSwitchAndSystrayIcon()

        Catch ex As Exception
        End Try

    End Sub


    Private Sub LoadListViewFromDatabase()

        Dim Source As New List(Of MyTable1ClassForTheListView)
        Source = LoadDatabaseIntoAList()

        Try

            For x As Integer = 0 To Source.Count - 1

                If Source.Item(x).Admin Then
                    Source.Item(x).DisplayableAdminText = gWordToUseToDenoteAdminInListView
                End If

                Source.Item(x).DisplayableStartingWindowStateText = Source.Item(x).StartingWindowState.ConvertStartingWindowStateToAString

            Next

            ListView1.ItemsSource = Nothing
            ListView1.Items.Clear()
            ListView1.ItemsSource = Source

        Catch ex As Exception
            Dim Result As MessageBoxResult = TopMostMessageBox(gCurrentOwner, ex.Message.ToString, "Push2Run - Error", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK, System.Windows.MessageBoxOptions.None)
        End Try

    End Sub


    Private Function LoadDatabaseIntoAList(Optional ByVal WithoutFiltering As Boolean = False) As List(Of MyTable1ClassForTheListView)

        Dim ReturnValue As New List(Of MyTable1ClassForTheListView)

        Try

            gBossLoadUnderway = True

            ReturnValue.Clear()

            Dim sSQL As String = "SELECT * FROM Table1 ORDER BY SortOrder ASC ;"
            Dim SQLiteConnect As New SQLiteConnection(gSQLiteConnectionString)
            Dim SQLiteCommand As SQLiteCommand = New SQLite.SQLiteCommand(sSQL, SQLiteConnect)

            SQLiteConnect.Open()

            Dim SQLiteDataReader As SQLiteDataReader = SQLiteCommand.ExecuteReader(CommandBehavior.CloseConnection)
            'Dim dsc As String

            Dim x As Int32 = 1

            While SQLiteDataReader.Read()

                Dim WorkingRecord As New MyTable1ClassForTheListView()

                With WorkingRecord

                    .ID = SQLiteDataReader.GetInt32(DatabaseColumns.ID)
                    .SortOrder = SQLiteDataReader.GetInt32(DatabaseColumns.SortOrder)
                    .DesiredStatus = SQLiteDataReader.GetInt32(DatabaseColumns.DesiredStatus)
                    .WorkingStatus = SQLiteDataReader.GetInt32(DatabaseColumns.WorkingStatus)

                    If (SQLiteDataReader.GetString(DatabaseColumns.Description) IsNot Nothing) AndAlso (SQLiteDataReader.GetString(DatabaseColumns.Description) <> String.Empty) Then
                        .Description = EncryptionClass.Decrypt(SQLiteDataReader.GetString(DatabaseColumns.Description))
                    End If

                    If (SQLiteDataReader.GetString(DatabaseColumns.ListenFor) IsNot Nothing) AndAlso (SQLiteDataReader.GetString(DatabaseColumns.ListenFor) <> String.Empty) Then
                        .ListenFor = EncryptionClass.Decrypt(SQLiteDataReader.GetString(DatabaseColumns.ListenFor))
                    End If

                    If (SQLiteDataReader.GetString(DatabaseColumns.Open) IsNot Nothing) AndAlso (SQLiteDataReader.GetString(DatabaseColumns.Open) <> String.Empty) Then
                        .Open = EncryptionClass.Decrypt(SQLiteDataReader.GetString(DatabaseColumns.Open))
                    End If

                    If (SQLiteDataReader.GetString(DatabaseColumns.Parameters) IsNot Nothing) AndAlso (SQLiteDataReader.GetString(DatabaseColumns.Parameters) <> String.Empty) Then
                        .Parameters = EncryptionClass.Decrypt(SQLiteDataReader.GetString(DatabaseColumns.Parameters))
                    End If

                    If (SQLiteDataReader.GetString(DatabaseColumns.StartIn) IsNot Nothing) AndAlso (SQLiteDataReader.GetString(DatabaseColumns.StartIn) <> String.Empty) Then
                        .StartIn = EncryptionClass.Decrypt(SQLiteDataReader.GetString(DatabaseColumns.StartIn))
                    End If

                    If (SQLiteDataReader.GetValue(DatabaseColumns.Admin) IsNot Nothing) Then
                        .Admin = SQLiteDataReader.GetValue(DatabaseColumns.Admin)
                    End If

                    If (SQLiteDataReader.GetValue(DatabaseColumns.StartingWindowState) IsNot Nothing) Then
                        .StartingWindowState = SQLiteDataReader.GetValue(DatabaseColumns.StartingWindowState)
                    End If

                    If (SQLiteDataReader.GetString(DatabaseColumns.KeysToSend) IsNot Nothing) AndAlso (SQLiteDataReader.GetString(DatabaseColumns.KeysToSend) <> String.Empty) Then
                        .KeysToSend = EncryptionClass.Decrypt(SQLiteDataReader.GetString(DatabaseColumns.KeysToSend))
                    End If

                    ' hide disabled cards v4.8.1

                    If .Description = gMasterSwitch Then
                    Else
                        If ViewDisabledCards Then
                        Else
                            If (.DesiredStatus = StatusValues.SwitchOff) OrElse (.WorkingStatus = StatusValues.SwitchOff) Then
                                GoTo SkipRecord
                            End If
                        End If
                    End If

                    ' filter logic here

                    If WithoutFiltering Then

                    Else

                        ' Filtering starts here

                        If FilterIsActive AndAlso (.Description = gMasterSwitch) Then
                            GoTo SkipRecord 'if any filtering is in play, then do not show the master switch
                        End If

                        If tbFilterDescription.Text.Trim > String.Empty Then
                            If .Description.ToLower.Contains(tbFilterDescription.Text.ToLower) Then
                            Else
                                GoTo SkipRecord
                            End If
                        End If

                        If tbFilterListenFor.Text.Trim > String.Empty Then
                            If .ListenFor.ToLower.Contains(tbFilterListenFor.Text.ToLower) Then
                            Else
                                GoTo SkipRecord
                            End If
                        End If

                        If tbFilterOpen.Text.Trim > String.Empty Then
                            If .Open.ToLower.Contains(tbFilterOpen.Text.ToLower) Then
                            Else
                                GoTo SkipRecord
                            End If
                        End If

                        If tbFilterStartIn.Text.Trim > String.Empty Then
                            If .StartIn.ToLower.Contains(tbFilterStartIn.Text.ToLower) Then
                            Else
                                GoTo SkipRecord
                            End If
                        End If

                        If tbFilterParameters.Text.Trim > String.Empty Then
                            If .Parameters.ToLower.Contains(tbFilterParameters.Text.ToLower) Then
                            Else
                                GoTo SkipRecord
                            End If
                        End If

                        If tbFilterAdmin.Text.Trim > String.Empty Then

                            If (tbFilterAdmin.Text.Trim.ToLower = "y") OrElse (tbFilterAdmin.Text.Trim.ToLower = "ye") OrElse (tbFilterAdmin.Text.Trim.ToLower = "yes") Then
                                If .Admin Then
                                Else
                                    GoTo SkipRecord
                                End If
                            End If

                            If (tbFilterAdmin.Text.Trim.ToLower = "n") OrElse (tbFilterAdmin.Text.Trim.ToLower = "no") Then
                                If .Admin Then
                                    GoTo SkipRecord
                                End If
                            End If

                        End If

                        If tbFilterStartingWindowState.Text.Trim > String.Empty Then

                            Dim wsFilter = tbFilterStartingWindowState.Text.Trim.ToLower
                            Dim wsValue = .StartingWindowState.ConvertStartingWindowStateToAString.ToLower

                            If wsValue.StartsWith(wsFilter) Then
                            Else
                                GoTo SkipRecord
                            End If

                        End If

                        If tbFilterKeysToSend.Text.Trim > String.Empty Then
                            If .KeysToSend.ToLower.Contains(tbFilterKeysToSend.Text.ToLower) Then
                            Else
                                GoTo SkipRecord
                            End If
                        End If

                        'record was not filtered out

                        ' Filtering ends here

                    End If

                    ReturnValue.Add(WorkingRecord)
                    x += 1

SkipRecord:

                End With

            End While

            SQLiteDataReader.Close()
            SQLiteConnect.Close()

            SQLiteConnect.Dispose()
            SQLiteCommand.Dispose()
            SQLiteDataReader = Nothing

            gBossLoadUnderway = False

        Catch ex As Exception

        End Try

        Return ReturnValue

    End Function


    Private Sub Systray_MenuPasswordRequired_click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Systray_MenuPasswordRequired.Click

        GetThePassword()

    End Sub


    Private Sub WindowBoss_StateChanged(sender As Object, e As EventArgs) Handles Me.StateChanged

        If Me.Systray_MenuShowBoss.Checked Then Exit Sub  ' this keeps the taskbar icon in place when the main window should be shown

        If gIgnorWindowStateChanges Then Exit Sub

        If Me.WindowState = System.Windows.WindowState.Minimized Then
            Me.Visibility = System.Windows.Visibility.Hidden
            SetPriority(ProcessPriorityClass.BelowNormal)
            Me.ShowInTaskbar = False
        Else
            Me.ShowInTaskbar = True
        End If

    End Sub



    Private Sub MenuShowSessionLog_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Systray_MenuShowSessionLog.Click

        Systray_MenuShowSessionLog.Checked = Not Systray_MenuShowSessionLog.Checked
        MenuViewSessionLog.IsChecked = Systray_MenuShowSessionLog.Checked
        ToggleTheSessionLogOpenAndClosed()

    End Sub


    Private Sub MenuShowOptions_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Systray_MenuShowOptions.Click

        Systray_MenuShowOptions.Checked = Not Systray_MenuShowOptions.Checked

        If Systray_MenuShowOptions.Checked Then
            OpenOptionsWindowAndResetWebsocketsIfNeeded()
        Else
            RaiseAnEventToCloseOptions()
        End If

    End Sub



    Private Sub MenuShowAboutHlep_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Systray_MenuShowAboutHelp.Click

        OpenOrCloseAboutHelp()

    End Sub


    Private Sub SysTrayIcon_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles SysTrayIcon.DoubleClick

        If gPasswordWasCorrectlyEnteredInPasswordWindow Then
            OpenMainWindow()
        Else
            GetThePassword()
            If gPasswordWasCorrectlyEnteredInPasswordWindow Then
                OpenMainWindow()
            End If
        End If

    End Sub

    Friend Sub OpenOptionsWindowAndResetWebsocketsIfNeeded()


        Dim HoldUsePushbullet As Boolean = My.Settings.UsePushbullet
        Dim HoldUsePushover As Boolean = My.Settings.UsePushover

        Dim HoldPushBulletAPIKey As String = My.Settings.PushBulletAPI

        Dim HoldPushoverUserID As String = My.Settings.PushoverUserID
        Dim HoldPushoverDeviceName As String = My.Settings.PushoverDeviceName

        OpenOptions()

        SetupForDropBoxProcessing()

        If My.Settings.UsePushbullet Then

            If My.Settings.PushBulletAPI = HoldPushBulletAPIKey Then
            Else
                're-open the websocket if the api key has changed
                Call OpenThePushbulletWebSocket()
            End If

        Else

            If HoldUsePushbullet Then
                ' settings were in place to use pushbullet but now they are not, so close the socket
                CloseThePushbulletWebSocket()
            End If

        End If

        If My.Settings.UsePushover Then

            If (My.Settings.PushoverUserID = HoldPushoverUserID) AndAlso (My.Settings.PushoverDeviceName = HoldPushoverDeviceName) Then
                DisablePushoverProcessing = True
                DeletePushoverMessages()
                DisablePushoverProcessing = False
            Else

                If IsPushOverFullyConfigured() Then
                    're-open the websocket if the userid or devicename have changed
                    DisablePushoverProcessing = True
                    OpenThePushoverWebSocket()
                    DeletePushoverMessages()
                    DisablePushoverProcessing = False
                End If

            End If

        Else

            If HoldUsePushover Then
                ' settings were in place to use pushover but now they are not, so close the socket
                CloseThePushoverWebSocket()
            End If

        End If

    End Sub

    Private Sub OpenOptions()

        Try

            Systray_MenuShowOptions.Checked = True
            Dim WindowOptions As WindowOptions = New WindowOptions

            gCurrentOwner = WindowOptions.Owner

            Dim hUseMQTT As Boolean = My.Settings.UseMQTT
            Dim hMQTTBroker As String = My.Settings.MQTTBroker
            Dim hMQTTPort As Integer = My.Settings.MQTTPort
            Dim hMQTTUser As String = My.Settings.MQTTUser
            Dim hPassword As String = My.Settings.MQTTPassword
            Dim gMQTTFilter As String = My.Settings.MQTTFilter

            WindowOptions.ShowDialog()
            WindowOptions = Nothing

            Try

                ' Special handling to deal with MQTT changes in Options

                If (My.Settings.UseMQTT = hUseMQTT) AndAlso (hMQTTBroker = My.Settings.MQTTBroker) AndAlso (hMQTTPort = My.Settings.MQTTPort) AndAlso (hMQTTUser = My.Settings.MQTTUser) AndAlso (hPassword = My.Settings.MQTTPassword) AndAlso (gMQTTFilter = My.Settings.MQTTFilter) Then
                    Exit Try  'nothing changed
                End If

                If (Not My.Settings.UseMQTT) AndAlso (Not hUseMQTT) Then
                    Exit Try ' Before or after the options were changed MQTT was/is not used
                End If

                If My.Settings.UseMQTT AndAlso (Not hUseMQTT) Then
                    'MQTT now needed
                    StartMQTTThreadWaitTillDone()
                    UpdateSubscriptionsWaitUntiDone()
                    Exit Try
                End If

                If (Not My.Settings.UseMQTT) AndAlso hUseMQTT Then
                    'MQTT no longer needed
                    MQTT_Disconnect().Wait()
                    Exit Try
                End If

                If (hMQTTPort <> My.Settings.MQTTPort) OrElse (hMQTTUser <> My.Settings.MQTTUser) OrElse (hPassword <> My.Settings.MQTTPassword) Then
                    MQTT_Disconnect().Wait()
                    StartMQTTThreadWaitTillDone()
                    UpdateSubscriptionsWaitUntiDone()
                    Exit Try
                End If

                If gMQTTFilter <> My.Settings.MQTTFilter Then
                    UpdateSubscriptionsWaitUntiDone()
                    Exit Try
                End If


            Catch ex As Exception

            End Try

            gCurrentOwner = Application.Current.MainWindow

            Systray_MenuShowOptions.Checked = False

            StartupShortCut("Ensure")

            MakeTopMost(SafeNativeMethods.FindWindow(Nothing, Me.Title), My.Settings.AlwaysOnTop)

            UpdateStatusBar()

        Catch ex As Exception

        End Try

    End Sub

    Private WindowSessionLog As WindowSessionLog

    Friend Sub UncheckSessionLogCheckbox()

        MenuViewSessionLog.IsChecked = False
        Systray_MenuShowSessionLog.Checked = False

    End Sub

    Private Sub ToggleTheSessionLogOpenAndClosed()

        If MenuViewSessionLog.IsChecked Then

            'open the session log

            If WindowSessionLog IsNot Nothing Then WindowSessionLog.Close()
            WindowSessionLog = Nothing
            WindowSessionLog = New WindowSessionLog
            WindowSessionLog.Show()

            If WindowSessionLog.WindowState = WindowState.Minimized Then
                WindowSessionLog.WindowState = WindowState.Normal
            End If

            WindowSessionLog.BringIntoView()

            SessionLogIsOpen = True
            MenuViewSessionLog.IsChecked = True

        Else

            'close the session log

            If WindowSessionLog IsNot Nothing Then
                WindowSessionLog.Close()
                WindowSessionLog = Nothing
            End If

            SessionLogIsOpen = False
            IndexToBeSelectedOnReload = -1

        End If

        Systray_MenuShowSessionLog.Checked = MenuViewSessionLog.IsChecked

    End Sub

    Private WindowAbout As WindowAbout = New WindowAbout

    Friend Sub UncheckAboutCheckbox()

        AboutHelpWindowIsOpen = False
        MenuViewAboutHelp.IsChecked = False
        Systray_MenuShowAboutHelp.Checked = False

    End Sub

    Private Sub OpenOrCloseAboutHelp()

        AboutHelpWindowIsOpen = Not AboutHelpWindowIsOpen

        MenuViewAboutHelp.IsChecked = AboutHelpWindowIsOpen
        Systray_MenuShowAboutHelp.Checked = AboutHelpWindowIsOpen

        If AboutHelpWindowIsOpen Then

            WindowAbout = New WindowAbout

            If Me.WindowState = WindowState.Normal Then
                WindowAbout.WindowStartupLocation = WindowStartupLocation.Manual
                WindowAbout.Top = Me.Top + 65
                WindowAbout.Left = Me.Left + 100
            End If

            WindowAbout.Show()

        Else

            If WindowAbout IsNot Nothing Then
                WindowAbout.Close()
                WindowAbout = Nothing
            End If

            MakeTopMost(SafeNativeMethods.FindWindow(Nothing, Me.Title), My.Settings.AlwaysOnTop)

        End If

    End Sub

    Delegate Sub OpenMainWindowCallback()

    Private Sub OpenMainWindow()

        gForceTheShowingOfTheMainWindowOnRestart = False

        DeterimineTheStateOfTheEnvironment() 'v4.6

        SetPriority(ProcessPriorityClass.Normal)

        If (Me.WindowState = System.Windows.WindowState.Minimized) OrElse (Me.Visibility = System.Windows.Visibility.Hidden) Then
            'all these cursor calls help the window open correctly
            SeCursor(CursorState.Wait)
            Me.Visibility = System.Windows.Visibility.Visible
            SeCursor(CursorState.Wait)
            Me.WindowState = System.Windows.WindowState.Normal
            SeCursor(CursorState.Wait)
        End If

        LoadWindowLocationSizeAndColumnWidths()

        Systray_MenuShowBoss.Checked = True

        Me.Show()

        Me.Activate()

        SeCursor(CursorState.Normal)

    End Sub


    Private Sub MenuPause_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Systray_MenuPause.CheckedChanged

        If gBossLoadUnderway Then Exit Sub

        ConfirmMasterSwitchAndSystrayIcon()

        If Systray_MenuPause.Checked Then
            WorkerTimer.Stop()
            Log("Master switch was turned off")
            Log("")
        Else
            WorkerTimer.Start()
            Log("Master switch was turned on")
            Log("")
        End If

    End Sub


    Private Sub ConfirmMasterSwitchAndSystrayIcon()

        If gPasswordWasCorrectlyEnteredInPasswordWindow AndAlso (Not Systray_MenuPause.Checked) Then
            TurnMasterSwitchOn(True)
        Else
            TurnMasterSwitchOn(False)
        End If

        ConfirmSystrayIcons()

    End Sub

    'Enum CurrentIcon
    '    Unknown = 0
    '    Normal = 1
    '    Problem = 2
    'End Enum

    Enum IconColour
        NotYetSet = 0
        Normal = 1
        Yellow = 2
        Red = 3
    End Enum


    Private Sub ConfirmSystrayIcons()

        Static CurrentIconColour As IconColour = IconColour.NotYetSet

        ' systray icon will be red if Push2Run is pause, otherwise
        ' it will be green if all enabled triggers are good
        ' it will be yellow if some selected triggers are good and others are not
        ' it will be red if all enabled triggers are not good

        ' can not determine if the Dropbox service is active, if it enabled assume it is

        ' Only update the system icon when needed

        If Systray_MenuPause.Checked Then

            If CurrentIconColour = IconColour.Red Then
            Else
                CurrentIconColour = IconColour.Red
                SysTrayIcon.Icon = My.Resources.Push2Run_r
            End If

            Exit Sub

        End If

        Dim TotalNumberOfEnabledTriggers As Integer = 0
        Dim TotalNumberOfTriggersThatAreGood As Integer = 0

        If My.Settings.UseMQTT Then TotalNumberOfEnabledTriggers = 1
        If My.Settings.UsePushover Then TotalNumberOfEnabledTriggers += 1
        If My.Settings.UsePushbullet Then TotalNumberOfEnabledTriggers += 1
        If My.Settings.UseDropbox Then TotalNumberOfEnabledTriggers += 1

        If My.Settings.UseMQTT AndAlso (gMQTTConnectionStatus = ConnectionStatus.Connected) Then TotalNumberOfTriggersThatAreGood = 1
        If My.Settings.UsePushover AndAlso (gPushoverConnectionStatus = ConnectionStatus.Connected) Then TotalNumberOfTriggersThatAreGood += 1
        If My.Settings.UsePushbullet AndAlso (gPushbulletConnectionStatus = ConnectionStatus.Connected) Then TotalNumberOfTriggersThatAreGood += 1
        If My.Settings.UseDropbox Then TotalNumberOfTriggersThatAreGood += 1

        Select Case TotalNumberOfTriggersThatAreGood

            Case Is = 0

                If CurrentIconColour <> IconColour.Red Then
                    CurrentIconColour = IconColour.Red
                    SysTrayIcon.Icon = My.Resources.Push2Run_r

                End If

            Case Is = TotalNumberOfEnabledTriggers

                If CurrentIconColour <> IconColour.Normal Then
                    CurrentIconColour = IconColour.Normal
                    SysTrayIcon.Icon = My.Resources.Push2Run
                End If

            Case Else

                If CurrentIconColour <> IconColour.Yellow Then
                    CurrentIconColour = IconColour.Yellow
                    SysTrayIcon.Icon = My.Resources.Push2Run_y
                End If

        End Select

    End Sub


    Private Sub MenuExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Systray_MenuExit.Click, MenuExit.Click

        Dim OKToExit As Boolean = True

        If My.Settings.ConfirmExit Then
            If TopMostMessageBox(gCurrentOwner, "Are you sure you want to exit?", "Push2Run - Exit confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No, System.Windows.MessageBoxOptions.None) = vbNo Then
                OKToExit = False
            End If
        End If

        If OKToExit Then
            OKToClose = True
            Me.Close()
        End If

    End Sub

    Private Sub Systray_MenuShowBoss_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Systray_MenuShowBoss.Click

        'this handles clicking on the red 'x' on the main window

        If Systray_MenuShowBoss.Checked Then
            OKToClose = False
            Me.Close()
        Else
            OpenMainWindow()
        End If

    End Sub



    Private Sub WindowBoss_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles Me.Closing

        If OKToClose Then

            gShutdownUnderway = True

            If gEarlyShutdown Then
            Else
                SaveWindowLocationSizeAndColumnWidthsAndOthers()
            End If

            SysTrayIcon.Visible = False

            DeleteAllTempFiles()

            WorkerTimer.Stop()
            WorkerTimer.Dispose()

            If PushbulletWebSocket IsNot Nothing Then
                If PushbulletWebSocket.State = PushbulletWebSocket.State.Open Then PushbulletWebSocket.Close()
            End If

            Try

                If PushbulletWebSocket IsNot Nothing Then
                    RemoveHandler PushbulletWebSocket.Opened, AddressOf Websocket_Opened_Pushbullet
                    RemoveHandler PushbulletWebSocket.MessageReceived, AddressOf Websocket_MessageReceived_Pushbullet
                    RemoveHandler PushbulletWebSocket.DataReceived, AddressOf Websocket_DataReceived_Pushbullet
                    RemoveHandler PushbulletWebSocket.Error, AddressOf Websocket_Error_Pushbullet
                    RemoveHandler PushbulletWebSocket.Closed, AddressOf Websocket_Closed_Pushbullet
                End If

            Catch ex As Exception

            End Try

            If PushoverWebSocket IsNot Nothing Then
                If PushoverWebSocket.State = PushoverWebSocket.State.Open Then PushoverWebSocket.Close()
            End If

            Try

                If PushoverWebSocket IsNot Nothing Then
                    RemoveHandler PushoverWebSocket.Opened, AddressOf Websocket_Opened_Pushover
                    RemoveHandler PushoverWebSocket.MessageReceived, AddressOf Websocket_MessageReceived_Pushover
                    RemoveHandler PushoverWebSocket.DataReceived, AddressOf Websocket_DataReceived_Pushover
                    RemoveHandler PushoverWebSocket.Error, AddressOf Websocket_Error_Pushover
                    RemoveHandler PushoverWebSocket.Closed, AddressOf Websocket_Closed_Pushover
                End If

            Catch ex As Exception

            End Try

            Try

                If gNetworkMonitoringOn Then
                    RemoveHandler NetworkChange.NetworkAddressChanged, AddressOf AddressChangedCallback
                    RemoveHandler NetworkChange.NetworkAvailabilityChanged, AddressOf AvailabilityChangedCallback
                End If

            Catch ex As Exception
            End Try

            Try

                If My.Settings.UseMQTT Then
                    MQTT_Disconnect().Wait()
                End If

            Catch ex As Exception
            End Try

            Log("")
            Log("Exited")
            Log("")

        Else

            SaveWindowLocationSizeAndColumnWidthsAndOthers() 'v1.8

            If My.Settings.ConfirmRedX Then
                Dim WindowClosingWarning = New WindowClosingWarning

                gCurrentOwner = WindowClosingWarning
                WindowClosingWarning.ShowDialog()
                gCurrentOwner = Application.Current.MainWindow

                WindowClosingWarning = Nothing
            End If

            Systray_MenuShowBoss.Checked = False

            Me.WindowState = System.Windows.WindowState.Minimized
            Me.Visibility = System.Windows.Visibility.Hidden

            SetPriority(ProcessPriorityClass.BelowNormal)

            e.Cancel = True

        End If

    End Sub

    Private Sub WindowBoss_Closed(sender As Object, e As EventArgs) Handles Me.Closed

        ConfirmMQTTConnectionTimer.Stop()
        ConfirmMQTTConnectionTimer.Dispose()

        ConfirmPushbulletConnectionTimer.Stop()
        ConfirmPushbulletConnectionTimer.Dispose()

        ConfirmPushoverConnectionTimer.Stop()
        ConfirmPushoverConnectionTimer.Dispose()

        WorkerTimer.Stop()
        WorkerTimer.Dispose()

        Try
            'todo: find why this is needed to complete shutdown
            System.Windows.Application.Current.ShutdownMode = ShutdownMode.OnExplicitShutdown
            System.Windows.Application.Current.Shutdown()
        Catch ex As Exception

        End Try

    End Sub


    Private Sub TurnMasterSwitchOn(ByVal TurnMasterOn As Boolean)

        If gBossLoadUnderway Then
        Else

            If (TurnMasterOn AndAlso (gMasterStatus = MonitorStatus.Running)) OrElse
               (Not TurnMasterOn) AndAlso (gMasterStatus = MonitorStatus.Stopped) Then
                Exit Sub
            End If
        End If

        If gIgnorWindowStateChanges Then
        Else
            AddToUndoTable(UndoRational.toggle_Master_Switch)
        End If

        SeCursor(CursorState.Wait)

        Try

            If TurnMasterOn Then

                gMasterStatus = MonitorStatus.Running

                'Turn the working status on for all records where the desired status is on 
                Dim WorkingRow As MyTable1ClassForTheListView

                WorkingRow = ListView1.Items(0)
                ChangeTheWorkingStatusSwitch(WorkingRow.ID, StatusValues.SwitchOn)

                For x As Int32 = 1 To ListView1.Items.Count - 1
                    WorkingRow = ListView1.Items(x)
                    If WorkingRow.DesiredStatus = StatusValues.SwitchOn Then
                        If WorkingRow.WorkingStatus = StatusValues.SwitchOff Then
                            ChangeTheWorkingStatusSwitch(WorkingRow.ID, StatusValues.SwitchOn)
                        End If
                    End If
                Next
                WorkingRow = Nothing
                LoadListViewFromDatabase()   ' A-b

            Else

                gMasterStatus = MonitorStatus.Stopped

                'Turn the working status off for all records
                Dim WorkingRow As MyTable1ClassForTheListView

                WorkingRow = ListView1.Items(0)

                ChangeTheWorkingStatusSwitch(WorkingRow.ID, StatusValues.SwitchOff)

                For x As Int32 = 1 To ListView1.Items.Count - 1
                    WorkingRow = ListView1.Items(x)
                    If WorkingRow.WorkingStatus = StatusValues.SwitchOn Then
                        ChangeTheWorkingStatusSwitch(WorkingRow.ID, StatusValues.SwitchOff)
                    End If
                Next
                WorkingRow = Nothing
                LoadListViewFromDatabase()

            End If

            ListView1.SelectedItem = ListView1.Items(0)
            ListView1.ScrollIntoView(ListView1.SelectedItem)

        Catch ex As Exception
        End Try

        SeCursor(CursorState.Normal)

    End Sub


    Private Sub ListView1_SelectionChanged(ByVal sender As Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles ListView1.SelectionChanged

        If gBossLoadUnderway Then Exit Sub

        If ListView1.SelectedItems.Count = 0 Then Exit Sub

        Dim SelectedRow As MyTable1ClassForTheListView = ListView1.SelectedItem

        On Error Resume Next

        With gCurrentlySelectedRow

            .ID = SelectedRow.ID
            .SortOrder = SelectedRow.SortOrder
            .DesiredStatus = SelectedRow.DesiredStatus
            .WorkingStatus = SelectedRow.WorkingStatus
            .Description = SelectedRow.Description
            .ListenFor = SelectedRow.ListenFor
            .Open = SelectedRow.Open
            .Parameters = SelectedRow.Parameters
            .StartIn = SelectedRow.StartIn
            .Admin = (SelectedRow.DisplayableAdminText = gWordToUseToDenoteAdminInListView)
            .StartingWindowState = SelectedRow.DisplayableStartingWindowStateText.ConvertStartingWindowStateToANumber
            .KeysToSend = SelectedRow.KeysToSend

        End With

        SetLookOfMenus()

    End Sub


    Private Sub MenuContext_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuContextAdd.Click,
                MenuContextCopy.Click, MenuContextEdit.Click, MenuContextDelete.Click, MenuContextInsertABlankLine.Click, MenuContextMoveUp.Click, MenuContextMoveDown.Click, MenuContextUndo.Click, MenuContextRun.Click,
                MenuAdd.Click, MenuCopy.Click, MenuEdit.Click, MenuDelete.Click, MenuInsertABlankLine.Click, MenuMoveUp.Click, MenuMoveDown.Click, MenuSort.Click, MenuUndo.Click,
                MenuMoveToTop.Click, MenuMoveToBottom.Click, MenuContextMoveToTop.Click, MenuContextMoveToBottom.Click,
                MenuSwitch.Click, MenuContextSwitch.Click, MenuOptions.Click, MenuElevate.Click, MenuRun.Click,
                MenuViewDescription.Click, MenuViewListenFor.Click, MenuViewOpen.Click, MenuViewParameters.Click, MenuViewStartIn.Click, MenuViewAdmin.Click, MenuViewStartingWindowState.Click, MenuViewKeysToSend.Click,
                MenuViewDisabledCards.Click, MenuViewFilters.Click, MenuViewSessionLog.Click, MenuViewAboutHelp.Click, MenuImport.Click, MenuExport.Click

        If gIgnoreAction Then
        Else
            PerformAction(sender.tag)
        End If

    End Sub


    Private Sub WindowBoss_PreviewKeyDown(sender As Object, e As Input.KeyEventArgs) Handles Me.PreviewKeyDown, ListView1.PreviewKeyDown

        SetLookOfMenus()

        If (Keyboard.Modifiers = ModifierKeys.Alt) OrElse (Keyboard.Modifiers = ModifierKeys.Shift) OrElse (Keyboard.Modifiers = ModifierKeys.Windows) Then

        Else

            Select Case e.Key

                Case Is = Key.F1
                    PerformAction("About/Help")
                    e.Handled = True

                Case Is = Key.F2
                    If MenuAdd.IsEnabled Then PerformAction("Add")
                    e.Handled = True

                Case Is = Key.F3
                    If MenuCopy.IsEnabled Then PerformAction("Copy")
                    e.Handled = True

                Case Is = Key.F4
                    If MenuEdit.IsEnabled Then PerformAction("Edit")
                    e.Handled = True

                Case Is = Key.F5
                    If MenuMoveToTop.IsEnabled Then PerformAction("Move to top")
                    e.Handled = True

                Case Is = Key.F6
                    If MenuMoveUp.IsEnabled Then PerformAction("Move up")
                    e.Handled = True

                Case Is = Key.F7
                    If MenuMoveDown.IsEnabled Then PerformAction("Move down")
                    e.Handled = True

                Case Is = Key.F8
                    If MenuMoveToBottom.IsEnabled Then PerformAction("Move to bottom")
                    e.Handled = True

                Case Is = Key.F9
                    If MenuSwitch.IsEnabled Then PerformAction("Switch")
                    e.Handled = True

                    'Case Is = 156 'Key.F10

                Case Is = Key.F12
                    If (MenuRun.IsEnabled) Then PerformAction("Run")
                    e.Handled = True

                Case Is = Key.Insert
                    If MenuContextInsertABlankLine.IsEnabled Then PerformAction("Insert a blank line")
                    e.Handled = True

                Case Is = 32 ' Delete
                    If MenuContextDelete.IsEnabled Then PerformAction("Delete")
                    e.Handled = True

                Case Is = Key.Z

                    If Keyboard.Modifiers = ModifierKeys.Control Then

                        PerformAction("Undo")
                        e.Handled = True

                    End If

                Case Is = Key.Down 'v3.5.3

                    Try
                        If ViewDisabledCards Then
                            ListView1.SelectedItem = ListView1.Items(ListView1.SelectedItem.ID)
                        Else

                            Dim startingPoint As Integer = ListView1.SelectedItem.SortOrder

                            Dim FoundCurrentItem As Boolean = False
                            For Each item In ListView1.Items

                                If FoundCurrentItem Then
                                    If item.WorkingStatus = StatusValues.SwitchOn Then
                                        ListView1.SelectedItem = item
                                        Exit For
                                    End If
                                End If

                                If item.SortOrder = startingPoint Then
                                    FoundCurrentItem = True
                                End If

                            Next

                        End If


                    Catch ex As Exception
                        ListView1.SelectedItem = ListView1.Items(ListView1.SelectedItem.ID - 1)
                    End Try
                    ListView1.ScrollIntoView(ListView1.SelectedItem)

                    e.Handled = True

                Case Is = Key.Up 'v3.5.3

                    Try

                        If ViewDisabledCards Then
                            ListView1.SelectedItem = ListView1.Items(ListView1.SelectedItem.ID - 2)
                        Else

                            Dim startingPoint As Integer = ListView1.SelectedItem.SortOrder

                            Dim priorItemThatIsOn As Object = Nothing
                            For Each item In ListView1.Items

                                If item.SortOrder = startingPoint Then

                                    If priorItemThatIsOn IsNot Nothing Then
                                        ListView1.SelectedItem = priorItemThatIsOn
                                        Exit For
                                    End If

                                End If

                                If item.WorkingStatus = StatusValues.SwitchOn Then
                                    priorItemThatIsOn = item
                                End If

                            Next

                        End If

                    Catch ex As Exception
                        ListView1.SelectedItem = ListView1.Items(ListView1.SelectedItem.ID - 1)
                    End Try

                    ListView1.ScrollIntoView(ListView1.SelectedItem)

                    e.Handled = True

                Case Else

                    Exit Sub

            End Select

        End If

        ' dn not put the e.Handled = True here as it cause problems

    End Sub

    Private Sub DoImport()

        Try
            AddToUndoTable(UndoRational.import)
            modImportAndExport.ImportDatabase(gImportFileName)
            RebuildTable1(MenuSort.IsChecked)
            ListView1.ScrollIntoView(ListView1.Items(1))
            RefreshListView(gGapBetweenSortIDsForDatabaseEntries)
        Catch ex As Exception
        End Try
        gImportFileName = String.Empty
    End Sub



    Private Sub PerformAction(ByVal ActionCode As String)

        SeCursor(CursorState.Wait)

        Try

            Select Case ActionCode

                Case Is = "Options"

                    OpenOptionsWindowAndResetWebsocketsIfNeeded()
                    Exit Try

                Case Is = "Import"

                    AddToUndoTable(UndoRational.import)
                    AutoCorrectTable1IfNeeded()
                    DoImport()
                    Exit Try


                Case Is = "Export"

                    ' note - you can't undo an export :-)

                    Dim MasterSwitchIsOn As Boolean = (Not Systray_MenuPause.Checked)

                    AutoCorrectTable1IfNeeded() 'v3.5

                    If IsAPasswordRequiredForBoss() Then

                        GetThePassword()

                        If gPasswordWasCorrectlyEnteredInPasswordWindow Then


                            modImportAndExport.ExportDatabase()

                        Else

                            'if the password is not know shut down

                            If MasterSwitchIsOn Then
                                TurnMasterSwitchOn(True)
                            End If
                            OKToClose = True
                            Me.Close()

                        End If
                    Else

                        modImportAndExport.ExportDatabase()

                    End If

                    Exit Try

                Case Is = "View Description"

                    ResetColumn(MenuViewDescription, ViewDescription, ListViewColumns.Description)
                    Exit Try

                Case Is = "View Listen for"

                    ResetColumn(MenuViewListenFor, ViewListenFor, ListViewColumns.ListenFor)
                    Exit Try

                Case Is = "View Open"

                    ResetColumn(MenuViewOpen, ViewOpen, ListViewColumns.Open)
                    Exit Try

                Case Is = "View Parameters"

                    ResetColumn(MenuViewParameters, ViewParameters, ListViewColumns.Parameters)
                    Exit Try

                Case Is = "View StartIn"

                    ResetColumn(MenuViewStartIn, ViewStartIn, ListViewColumns.StartIn)
                    Exit Try

                Case Is = "View Admin"

                    ResetColumn(MenuViewAdmin, ViewAdmin, ListViewColumns.Admin)
                    Exit Try

                Case Is = "View Window state"

                    ResetColumn(MenuViewStartingWindowState, ViewStartingWindowState, ListViewColumns.StartingWindowState)
                    Exit Try

                Case Is = "View Keys to send"

                    ResetColumn(MenuViewKeysToSend, ViewKeysToSend, ListViewColumns.KeysToSend)
                    Exit Try

                Case Is = "View Disabled Cards"

                    ShowOrHideDisabledCards()
                    Exit Try

                Case Is = "View Filters"

                    ShowOrHideFilters()
                    Exit Try

                Case Is = "Session log"

                    ToggleTheSessionLogOpenAndClosed()
                    Exit Try

                Case Is = "Run"

                    RunCurrentlySelectedRecord()
                    Exit Try

                Case Is = "ChangeAdminRights"

                    Try

                        'v4.6
                        If gRunningInASandbox Then

                            Dim Result As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "Cannot change privileges as program is running in a sandbox", "Push2Run", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK, System.Windows.MessageBoxOptions.None)

                        Else



                            If (Not gIsWindows10OrAbove) AndAlso (Not gIsUACOn) Then
                                ' special case, leave everything alone to prevent looping
                                Exit Try
                            End If

                            SaveWindowLocationSizeAndColumnWidthsAndOthers()

                            Dim sArgument As String

                            If gIsAdministrator Then
                                sArgument = "RestartNormal"
                            Else
                                sArgument = "RestartAdmin"
                            End If

                            ReloadPush2Run(sArgument)

                        End If

                    Catch ex As Exception
                        MsgBox(ex.ToString,, "Push2Run - Change Admin Rights")
                    End Try

                    Exit Try

                Case Is = "ElevateAtStartup"

                    Try

                        Me.Hide()
                        ReloadPush2Run("StartAdmin")

                    Catch ex As Exception
                        MsgBox(ex.ToString,, "Push2Run - ElevateAtStartup")
                    End Try

                    Exit Try

                Case Is = "About/Help"

                    OpenOrCloseAboutHelp()
                    Exit Try

            End Select


            If ActionCode = "Undo" Then
            Else
                If ListView1.SelectedItems.Count = 0 Then

                    If ActionCode = "Add" Then

                    Else

                        If gDropInProgress Then

                        Else

                            Exit Try

                        End If

                    End If

                End If
            End If

            Dim NewSortOrderPosition As Integer = gCurrentlySelectedRow.SortOrder

            If gMasterStatus = MonitorStatus.Running Then
                gMasterStatus = StatusValues.SwitchOn
            Else
                gMasterStatus = StatusValues.SwitchOff
            End If

            Dim RebuildRequired As Boolean = False

            Select Case ActionCode

                Case Is = "Add"

                    AddToUndoTable(UndoRational.add)

                    If ListView1.SelectedIndex = -1 Then
                        ListView1.SelectedIndex = ListView1.Items.Count
                    End If

                    gCurrentlySelectedRow.WorkingStatus = StatusValues.SwitchOff
                    If GetAddChangeInfo("Add") Then

                        With gCurrentlySelectedRow
                            RebuildRequired = UpdateSortOrder(.SortOrder)
                            InsertARecord(.SortOrder, StatusValues.SwitchOff, StatusValues.SwitchOff, .Description, .ListenFor, .Open, .Parameters, .StartIn, .Admin, .StartingWindowState, .KeysToSend)
                            NewSortOrderPosition = .SortOrder
                        End With

                    Else
                        CancelAddToUndoTable()
                    End If


                Case Is = "Copy"

                    With gCurrentlySelectedRow

                        If .SortOrder <> MasterControlSwitchID Then

                            AddToUndoTable(UndoRational.copy)

                            RebuildRequired = UpdateSortOrder(.SortOrder)

                            If (.Description = "") AndAlso (.ListenFor = "") AndAlso (.Open = "") AndAlso (.Parameters = "") AndAlso (.StartIn = "") AndAlso (.Admin = False) Then
                                InsertARecord(.SortOrder, StatusValues.NoSwitch, StatusValues.NoSwitch, "", "", "", "", "", False, 0, "")
                            Else
                                InsertARecord(.SortOrder, .DesiredStatus, .WorkingStatus, .Description & " - copy", .ListenFor, .Open, .Parameters, .StartIn, .Admin, .StartingWindowState, .KeysToSend)
                            End If

                            NewSortOrderPosition = .SortOrder

                        End If

                    End With

                Case Is = "Edit"

                    ChangeTheSelectedRow()
                    RebuildRequired = False

                Case Is = "Delete"

                    If gCurrentlySelectedRow.SortOrder = MasterControlSwitchID Then
                        Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "The Master Switch can't be deleted", "Push2Run - Info", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK, System.Windows.MessageBoxOptions.None)
                        Exit Try
                    End If

                    Dim OKToDelete As Boolean = True

                    If My.Settings.ConfirmDelete Then
                        If TopMostMessageBox(gCurrentOwner, "Are you sure you want to delete this entry?", "Push2Run - Delete confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No, System.Windows.MessageBoxOptions.None) = vbNo Then
                            OKToDelete = False
                        End If
                    End If

                    If OKToDelete Then

                        AddToUndoTable(UndoRational.delete)

                        Dim NewRecordToBePositioned As MyTable1Class

                        Dim NextRecordIDOffsetFromCurrentRecord As Integer = 1
                        NewRecordToBePositioned = ReadARecord(gCurrentlySelectedRow.ID + NextRecordIDOffsetFromCurrentRecord)

                        'look forward to find the next record to be positioned at (code handles if some records were deleted)

                        Dim LastIDInTable As Integer = GetLastIDInDatabase()

                        If (NewRecordToBePositioned Is Nothing) OrElse (NewRecordToBePositioned.ID = 0) Then

                            If LastIDInTable > gCurrentlySelectedRow.ID Then

                                ' continue to look forward to find the next ID (code below covers case some records have been deleted since the last reload)
                                While (NewRecordToBePositioned Is Nothing) OrElse (NewRecordToBePositioned.ID = 0)
                                    NextRecordIDOffsetFromCurrentRecord += 1
                                    NewRecordToBePositioned = ReadARecord(gCurrentlySelectedRow.ID + NextRecordIDOffsetFromCurrentRecord)
                                End While

                            Else

                                ' start look backward find the prior ID (code below covers case some records have been deleted since the last reload)
                                ' eventually will hit the master switch record if there are no other records left
                                NextRecordIDOffsetFromCurrentRecord = 0
                                While (NewRecordToBePositioned Is Nothing) OrElse (NewRecordToBePositioned.ID = 0)
                                    NextRecordIDOffsetFromCurrentRecord -= 1
                                    NewRecordToBePositioned = ReadARecord(gCurrentlySelectedRow.ID + NextRecordIDOffsetFromCurrentRecord)
                                End While

                            End If

                        End If

                        DeleteARecord(gCurrentlySelectedRow.ID)
                        NewSortOrderPosition = NewRecordToBePositioned.SortOrder

                        RebuildRequired = False

                    End If

                Case Is = "Insert a blank line"

                    AddToUndoTable(UndoRational.insert_a_blank_line)
                    RebuildRequired = UpdateSortOrder(gCurrentlySelectedRow.SortOrder)
                    InsertARecord(gCurrentlySelectedRow.SortOrder, StatusValues.NoSwitch, StatusValues.NoSwitch, "", "", "", "", "", False, 0, "")
                    NewSortOrderPosition = gCurrentlySelectedRow.SortOrder

                Case Is = "Move up" ' 

                    If ListView1.Items.Count > 2 Then

                        Dim PriorID, PriorSortOrder As Integer

                        'Find of row above the currently selected row and get it's sort id

                        For x As Int32 = 1 To ListView1.Items.Count - 1
                            Dim aRow As MyTable1ClassForTheListView = ListView1.Items(x)
                            If aRow.ID = gCurrentlySelectedRow.ID Then
                                PriorID = ListView1.Items(x - 1).ID
                                PriorSortOrder = ListView1.Items(x - 1).SortOrder
                                Exit For
                            End If
                        Next

                        AddToUndoTable(UndoRational.move_up)

                        If SwapTwoRecordsOnTheDatabase(gCurrentlySelectedRow.ID, PriorID) Then
                            NewSortOrderPosition = PriorSortOrder
                        Else
                            NewSortOrderPosition = gCurrentlySelectedRow.SortOrder
                        End If

                    End If


                    RebuildRequired = False

                Case Is = "Move down"

                    Dim MoveDownIsPossible As Boolean = False
                    Dim NextID, NextSortOrder As Integer

                    If ListView1.Items.Count > 2 Then

                        'Find of row below the currently selected row and get it's sort id

                        For x As Int32 = 1 To ListView1.Items.Count - 2

                            Dim aRow As MyTable1ClassForTheListView = ListView1.Items(x)

                            If aRow.ID = gCurrentlySelectedRow.ID Then

                                NextID = ListView1.Items(x + 1).ID
                                NextSortOrder = ListView1.Items(x + 1).SortOrder
                                MoveDownIsPossible = True
                                Exit For

                            End If

                        Next

                    End If

                    If MoveDownIsPossible Then

                        AddToUndoTable(UndoRational.move_down)

                        If SwapTwoRecordsOnTheDatabase(gCurrentlySelectedRow.ID, NextID) Then
                            NewSortOrderPosition = NextSortOrder
                        Else
                            NewSortOrderPosition = gCurrentlySelectedRow.SortOrder
                        End If

                    End If

                    RebuildRequired = False

                Case Is = "Move to top"

                    AddToUndoTable(UndoRational.move_top)
                    ChangeTheSortOrderOfARecord(gCurrentlySelectedRow.ID, 2) '  
                    RebuildRequired = True

                Case Is = "Move to bottom"

                    AddToUndoTable(UndoRational.move_bottom)
                    ChangeTheSortOrderOfARecord(gCurrentlySelectedRow.ID, Integer.MaxValue)
                    RebuildRequired = True

                Case Is = "Sort"

                    gMenuSort = MenuSort.IsChecked 'v 2.5.3
                    If MenuSort.IsChecked Then
                        AddToUndoTable(UndoRational.sort)
                        RebuildRequired = True
                    Else
                        AddToUndoTable(UndoRational.remove_sort)
                    End If

                Case Is = "Switch"

                    AddToUndoTable(UndoRational.toggle_a_switch)
                    Select Case gCurrentlySelectedRow.WorkingStatus

                        Case Is = StatusValues.SwitchOn, StatusValues.SwitchOff
                            AddToUndoTable(UndoRational.toggle_a_switch)
                            ToggleSwitch()

                    End Select
                    RebuildRequired = False

                Case Is = "Undo"

                    If MyUndoTableIndex = 0 Then

                        Beep()

                    Else

                        UndoDisplayTimer.Start()
                        UpdateUndoDisplayMessage("Undo " & MyUndoTable(MyUndoTableIndex).Rational.ToString.Replace("_", " ") & " - complete!")

                        If (MyUndoTable(MyUndoTableIndex).Rational = UndoRational.sort) OrElse (MyUndoTable(MyUndoTableIndex).Rational = UndoRational.remove_sort) Then
                            gIgnoreAction = True
                            MenuSort.IsChecked = Not MenuSort.IsChecked
                            gIgnoreAction = False
                        End If

                        NewSortOrderPosition = RestoreFromUndoTable()

                    End If

                    RebuildRequired = False

            End Select

            If RebuildRequired Then
                RebuildTable1(MenuSort.IsChecked)
            End If

            If ActionCode = "Move to bottom" Then
                ListView1.ScrollIntoView(ListView1.Items(1))
            End If

            If ActionCode = "Move to top" Then
                NewSortOrderPosition = gGapBetweenSortIDsForDatabaseEntries
            End If

            RefreshListView(NewSortOrderPosition)
            'UpdateTheListView() '***************************

            Select Case ActionCode

                'reposition currently selected item
                Case Is = "Move to top"

                    ListView1.SelectedItem = ListView1.Items(1)
                    ListView1.ScrollIntoView(ListView1.SelectedItem)

                Case Is = "Move to bottom"

                    ListView1.SelectedItem = ListView1.Items(ListView1.Items.Count - 1)
                    ListView1.ScrollIntoView(ListView1.SelectedItem)

            End Select

        Catch ex As Exception

        End Try

        SeCursor(CursorState.Normal)

    End Sub

    Private Function UpdateSortOrder(ByRef SortOrderValue As Integer) As Boolean

        Dim Remainder As Integer = SortOrderValue Mod gGapBetweenSortIDsForDatabaseEntries

        Dim NextHalfPoint As Integer = gGapBetweenSortIDsForDatabaseEntries / 2

        While Remainder <> 0
            Remainder = Remainder Mod NextHalfPoint
            NextHalfPoint /= 2
        End While

        SortOrderValue += NextHalfPoint

        Return (NextHalfPoint <= 2)

    End Function

    Private Sub ReloadPush2Run(ByVal ArgumentsIn As String, Optional ByVal VerbIn As String = "runas")

        Try

            Dim myProcess As New Process

            With myProcess.StartInfo

                .CreateNoWindow = True
                .UseShellExecute = True

                'v4.6 for testing in debug mode (esp. with UAC off and running as admin)
#If DEBUG Then
                .WorkingDirectory = "C:\Program Files\Push2Run"
#Else
                .WorkingDirectory = System.Environment.CurrentDirectory
#End If

                .FileName = "Push2RunReloader.exe"
                .Verb = VerbIn
                .Arguments = ArgumentsIn

            End With

            Dim ProcessStarted As Boolean = myProcess.Start()

            Thread.Sleep(2000) ' added in v3.5

            If ProcessStarted Then

                OKToClose = True
                Me.Close()

                Exit Sub

            Else

                Dim Result As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "Function did not work as expected - " & ArgumentsIn, "Push2Run - Warning", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK, System.Windows.MessageBoxOptions.None)

            End If

        Catch ex As Exception

            'v4.6
            If ex.ToString.Contains("The operation was canceled by the user") Then
            Else
                MsgBox("Gnats:" & vbCrLf & ex.ToString)
            End If

        End Try

    End Sub
    Private Sub ResetColumn(ByVal MenuViewItem As Controls.MenuItem, ByRef ViewItem As Boolean, ByVal lvColumnNumber As Integer)

        EnsureAtLeastOneColumnIsVisible(MenuViewItem)

        ViewItem = MenuViewItem.IsChecked
        Dim gv As GridView = ListView1.View

        If ViewItem Then
            gv.Columns.Item(lvColumnNumber).Width = Double.NaN
        Else
            gv.Columns.Item(lvColumnNumber).Width = 0
        End If

        SetFiltersVisibility()

    End Sub

    Private Sub ShowOrHideDisabledCards()

        ViewDisabledCards = MenuViewDisabledCards.IsChecked

        RefreshListView(CurrentSortOrder)

    End Sub

    Private Sub ShowOrHideFilters()

        ViewFilters = MenuViewFilters.IsChecked

        If ViewFilters Then

            'add filters in

            FilterBoundry.Visibility = Visibility.Visible
            FilterCanvas.Visibility = Visibility.Visible
            'the following lines force a redraw of the list view and return it to a state where its height is stretchable
            ListView1.Margin = New Thickness(ListView1.Margin.Left, ListView1.Margin.Top, ListView1.Margin.Right, ListView1.Margin.Bottom + FilterCanvas.ActualHeight)
            ListView1.Height = ListView1.ActualHeight - FilterCanvas.ActualHeight
            ListView1.Height = Double.NaN
            RefreshListView(CurrentSortOrder)

        Else

            ' remove filters

            ClearAllFilters()
            FilterBoundry.Visibility = Visibility.Hidden
            FilterCanvas.Visibility = Visibility.Hidden
            ListView1.Margin = New Thickness(ListView1.Margin.Left, ListView1.Margin.Top, ListView1.Margin.Right, ListView1.Margin.Bottom - FilterCanvas.ActualHeight)
            ListView1.Height = ListView1.ActualHeight + FilterCanvas.ActualHeight
            ListView1.Height = Double.NaN
            RefreshListView(CurrentSortOrder)

        End If

    End Sub

    Private Sub EnsureAtLeastOneColumnIsVisible(ByRef MenuItem As Controls.MenuItem)

        If MenuViewDescription.IsChecked OrElse MenuViewListenFor.IsChecked OrElse MenuViewOpen.IsChecked OrElse MenuViewParameters.IsChecked OrElse MenuViewStartIn.IsChecked OrElse MenuViewAdmin.IsChecked OrElse
                                                MenuViewStartingWindowState.IsChecked OrElse MenuViewKeysToSend.IsChecked Then
        Else
            Dim Result As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "All fileds cannot be unchecked at the same time.", "Push2Run - Warning", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK, System.Windows.MessageBoxOptions.None)
            MenuItem.IsChecked = Not MenuItem.IsChecked
        End If

    End Sub


    Private Function SwapTwoRecordsOnTheDatabase(ByVal RecordOneID As Integer, ByVal RecordTwoID As Integer) As Boolean

        Dim ReturnValue As Boolean = False

        Try

            If (RecordOneID = 1) OrElse (RecordTwoID = 1) Then

            Else

                Dim RecordOne, RecordTwo, RecordHold As MyTable1Class

                RecordOne = ReadARecord(RecordOneID)
                RecordTwo = ReadARecord(RecordTwoID)
                RecordHold = RecordTwo

                gBossLoadUnderway = True

                With RecordOne
                    ChangeARecord(.ID, RecordTwo.SortOrder, .DesiredStatus, .WorkingStatus, .Description, .ListenFor, .Open, .Parameters, .StartIn, .Admin, .StartingWindowState, .KeysToSend)
                End With

                With RecordHold
                    ChangeARecord(.ID, RecordOne.SortOrder, .DesiredStatus, .WorkingStatus, .Description, .ListenFor, .Open, .Parameters, .StartIn, .Admin, .StartingWindowState, .KeysToSend)
                End With

                gBossLoadUnderway = False

                RecordOne = Nothing
                RecordTwo = Nothing
                RecordHold = Nothing

                ReturnValue = True

            End If

        Catch ex As Exception

        End Try

        Return ReturnValue

    End Function


    Private Sub ListView1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles ListView1.MouseDoubleClick

        ' Double click on an entry to change it

        If ListView1.SelectedItems.Count = 0 Then Exit Sub 'don't react unless an entry has been selected

        If gCurrentlySelectedRow.DesiredStatus = StatusValues.NoSwitch Then Exit Sub 'don't react to blank lines

        Dim XPos As Double = e.GetPosition(relativeTo:=ListView1).X
        If XPos < 76 Then Exit Sub 'Don't react when switch column is clicked

        'v 3.4.2 exclude the master control record from being double clicked on to open
        If gCurrentlySelectedRow.Description = "Master Switch" Then Exit Sub

        Dim YPos As Double = e.GetPosition(relativeTo:=ListView1).Y
        If YPos < 21 Then Exit Sub ' exclude header row 
        ' If YPos > (ListView1.Items.Count * 21 + 30) Then Exit Sub ' exclude empty space on the window below the last entry

        Dim NewSortOrderPosition As Integer = gCurrentlySelectedRow.SortOrder

        PerformAction("Edit")

        RefreshListView(NewSortOrderPosition)

    End Sub

    Private EditWasCancelled As Boolean

    Private Sub ChangeTheSelectedRow()

        Dim OriginalWorkingStatusIsOn As Boolean = (gCurrentlySelectedRow.WorkingStatus = StatusValues.SwitchOn)

        EditWasCancelled = False

        AddToUndoTable(UndoRational.change_to_existing_Push2Run_card)

        'Turn Off Monitoring while in edit mode

        With gCurrentlySelectedRow

            If OriginalWorkingStatusIsOn Then
                .DesiredStatus = StatusValues.SwitchOff
                .WorkingStatus = StatusValues.SwitchOff
                ChangeARecord(.ID, .SortOrder, .DesiredStatus, .WorkingStatus, .Description, .ListenFor, .Open, .Parameters, .StartIn, .Admin, .StartingWindowState, .KeysToSend)
                RefreshListView(.SortOrder)
                LoadFromDatabase(.ID)
            End If

            'Edit the entry
            If GetAddChangeInfo("Edit") AndAlso (gReturnFromAddChangeDataChanged OrElse gDropInProgress) Then
                ChangeARecord(.ID, .SortOrder, StatusValues.SwitchOff, StatusValues.SwitchOff, .Description, .ListenFor, .Open, .Parameters, .StartIn, .Admin, .StartingWindowState, .KeysToSend)
                .DesiredStatus = StatusValues.SwitchOff
                .WorkingStatus = StatusValues.SwitchOff
            Else
                EditWasCancelled = True
                CancelAddToUndoTable()
            End If

            If OriginalWorkingStatusIsOn Then
                .DesiredStatus = StatusValues.SwitchOn
                .WorkingStatus = StatusValues.SwitchOn
                ChangeARecord(.ID, .SortOrder, .DesiredStatus, .WorkingStatus, .Description, .ListenFor, .Open, .Parameters, .StartIn, .Admin, .StartingWindowState, .KeysToSend)
                RefreshListView(.SortOrder)
                LoadFromDatabase(.ID)
            End If

        End With

    End Sub


    Private Sub ListView1_MouseLeftButtonUporDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles ListView1.MouseLeftButtonUp, ListView1.MouseRightButtonUp

        'This code toggles the on/off switches

        If ListView1.SelectedItems.Count = 0 Then Exit Sub 'only react if an entry has been selected

        If gCurrentlySelectedRow.DesiredStatus = StatusValues.NoSwitch Then Exit Sub 'don't react to blank lines

        Dim XPos As Double = e.GetPosition(relativeTo:=ListView1).X
        If (XPos < 10) OrElse (XPos > 70) Then Exit Sub 'only react when switch column is clicked

        ToggleSwitch()

        RefreshListView(gCurrentlySelectedRow.SortOrder)

    End Sub


    Private Sub ToggleSwitch()

        Dim SelectedRow As MyTable1ClassForTheListView = ListView1.SelectedItem
        Dim SelectedRow_ID As String = SelectedRow.ID

        'for all switches other than the master switch, only react if master switch is turned on
        If SelectedRow_ID > 1 Then
            If gMasterStatus = MonitorStatus.Stopped Then
                Exit Sub
            End If
        End If

        SeCursor(CursorState.Wait)

        Dim NewStatus As StatusValues

        Select Case gCurrentlySelectedRow.WorkingStatus

            Case Is = StatusValues.SwitchOn
                NewStatus = StatusValues.SwitchOff

            Case Is = StatusValues.SwitchOff
                NewStatus = StatusValues.SwitchOn

        End Select

        ChangeTheDesiredStatusSwitch(SelectedRow_ID, NewStatus)
        ChangeTheWorkingStatusSwitch(SelectedRow_ID, NewStatus)

        If SelectedRow_ID = MasterControlSwitchID Then

            'Master Control Switch Logic

            If NewStatus = StatusValues.SwitchOn Then
                Systray_MenuPause.Checked = False
                Systray_MenuPause.Checked = False
            Else
                Systray_MenuPause.Checked = True
            End If

            TurnMasterSwitchOn(Not Systray_MenuPause.Checked)

        Else

            LoadFromDatabase(SelectedRow_ID)

        End If

        SeCursor(CursorState.Normal)

    End Sub


    Private Sub MenuActions_SubmenuOpened(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuActions.SubmenuOpened, MenuView.SubmenuOpened, MenuHelp.SubmenuOpened

        SetLookOfMenus()

    End Sub

    Private Sub ListView1_ContextMenuOpening(ByVal sender As Object, ByVal e As System.Windows.Controls.ContextMenuEventArgs) Handles ListView1.ContextMenuOpening

        SetLookOfMenus()

    End Sub
    Private Sub SetLookOfMenus()

        Dim SortOrderOfFirstRow As Integer = 0
        Dim SortOrderOfSeconddRow As Integer = 0
        Dim SortOrderOfLastRow As Integer = 0

        Try ' v3.5.2


            If (ListView1.Items.Count <= 1) Then  'v3.5.2

            ElseIf (ListView1.Items.Count = 2) Then
                Dim lvi As MyTable1ClassForTheListView
                lvi = ListView1.Items(1)
                SortOrderOfSeconddRow = lvi.SortOrder
                SortOrderOfLastRow = SortOrderOfSeconddRow
                lvi = Nothing

            Else

                Dim lvi As MyTable1ClassForTheListView
                lvi = ListView1.Items(1)
                SortOrderOfSeconddRow = lvi.SortOrder
                lvi = ListView1.Items(ListView1.Items.Count - 1)
                SortOrderOfLastRow = lvi.SortOrder
                lvi = Nothing

            End If

        Catch ex As Exception

            ' no match found

        End Try


        If (ListView1.SelectedItems.Count = 0) OrElse ((ListView1.SelectedIndex = 0) AndAlso (Not FilterIsActive)) Then

            'Set Menu Items where no entry is selected or it is the master switch

            MenuAdd.IsEnabled = True
            MenuEdit.IsEnabled = False
            MenuCopy.IsEnabled = False
            MenuDelete.IsEnabled = False

            MenuMoveUp.IsEnabled = False
            MenuMoveDown.IsEnabled = False
            MenuInsertABlankLine.IsEnabled = False

            MenuMoveToTop.IsEnabled = False
            MenuMoveToBottom.IsEnabled = False

            MenuSwitch.IsEnabled = True

            Seperator5a.IsEnabled = False

            MenuRun.IsEnabled = False
            Seperator6a.Visibility = Visibility.Collapsed
            MenuRun.Visibility = Visibility.Collapsed

            Seperator4b.Visibility = Visibility.Collapsed
            MenuContextRun.Visibility = Visibility.Collapsed

        Else

            'Set Menu Items where an entry is selected

            MenuAdd.IsEnabled = True
            MenuEdit.IsEnabled = True
            MenuCopy.IsEnabled = True
            MenuDelete.IsEnabled = True

            MenuMoveUp.IsEnabled = Not MenuSort.IsChecked
            MenuMoveDown.IsEnabled = Not MenuSort.IsChecked
            MenuInsertABlankLine.IsEnabled = Not MenuSort.IsChecked

            MenuSwitch.IsEnabled = True

            Select Case gCurrentlySelectedRow.SortOrder

                Case Is = SortOrderOfFirstRow ' the is the master row
                    MenuEdit.IsEnabled = False
                    MenuCopy.IsEnabled = False
                    MenuDelete.IsEnabled = False
                    MenuMoveUp.IsEnabled = False
                    MenuMoveDown.IsEnabled = False
                    'Log("first")

                Case Is = SortOrderOfSeconddRow   ' this is the second row
                    MenuEdit.IsEnabled = True
                    MenuCopy.IsEnabled = True
                    MenuDelete.IsEnabled = True
                    MenuMoveUp.IsEnabled = False
                    ' this handles the case that the first row of data is also the last row of data
                    MenuMoveDown.IsEnabled = (gCurrentlySelectedRow.SortOrder <> SortOrderOfLastRow) AndAlso Not MenuSort.IsChecked

                Case Is = SortOrderOfLastRow
                    MenuEdit.IsEnabled = True
                    MenuCopy.IsEnabled = True
                    MenuDelete.IsEnabled = True
                    MenuMoveUp.IsEnabled = True AndAlso Not MenuSort.IsChecked
                    MenuMoveDown.IsEnabled = False

                Case Else
                    MenuEdit.IsEnabled = True
                    MenuCopy.IsEnabled = True
                    MenuDelete.IsEnabled = True
                    MenuMoveUp.IsEnabled = True AndAlso Not MenuSort.IsChecked
                    MenuMoveDown.IsEnabled = True AndAlso Not MenuSort.IsChecked

            End Select

            MenuMoveToTop.IsEnabled = MenuMoveUp.IsEnabled AndAlso Not MenuSort.IsChecked
            MenuMoveToBottom.IsEnabled = MenuMoveDown.IsEnabled AndAlso Not MenuSort.IsChecked

            If gCurrentlySelectedRow.ID > 1 Then

                Select Case gCurrentlySelectedRow.WorkingStatus

                    Case Is = StatusValues.SwitchOn, StatusValues.SwitchOff
                        MenuSwitch.IsEnabled = True
                    Case Else
                        MenuSwitch.IsEnabled = False

                End Select

            End If

            'Adjust for blank lines and where there is no description / program

            If (gCurrentlySelectedRow.DesiredStatus = StatusValues.NoSwitch) Then

                MenuEdit.IsEnabled = False

                MenuRun.IsEnabled = False

                Seperator6a.Visibility = Visibility.Collapsed
                MenuRun.Visibility = Visibility.Collapsed

                Seperator4b.Visibility = Visibility.Collapsed
                MenuContextRun.Visibility = Visibility.Collapsed

            Else

                If (gCurrentlySelectedRow.Description.Length > 0) AndAlso (gCurrentlySelectedRow.Open.Length > 0) Then
                    MenuRun.IsEnabled = True
                    MenuRun.Visibility = Visibility.Visible
                    MenuRun.Header = "_Run " & gCurrentlySelectedRow.Description

                ElseIf gCurrentlySelectedRow.Open.Length > 0 Then
                    MenuRun.IsEnabled = True

                    MenuRun.Visibility = Visibility.Visible
                    MenuRun.Header = "_Run currently selected entry"

                Else
                    MenuRun.IsEnabled = False
                    MenuRun.Visibility = Visibility.Collapsed

                End If

                Seperator6a.Visibility = MenuRun.Visibility
                Seperator4b.Visibility = MenuRun.Visibility
                MenuContextRun.Visibility = MenuRun.Visibility

            End If

        End If

        If MenuSwitch.IsEnabled Then

            If gCurrentlySelectedRow.WorkingStatus = StatusValues.SwitchOn Then
                MenuSwitch.Header = "_Switch Off"
            Else
                MenuSwitch.Header = "_Switch On"
            End If

        Else

            MenuSwitch.Header = "_Switch On/Off"

        End If

        MenuUndo.IsEnabled = (MyUndoTableIndex > 0)

        If FilterIsActive Then 'v2.0.2

            MenuAdd.IsEnabled = False
            MenuCopy.IsEnabled = False
            MenuDelete.IsEnabled = False

            MenuMoveUp.IsEnabled = False
            MenuMoveDown.IsEnabled = False
            MenuInsertABlankLine.IsEnabled = False

            MenuMoveToTop.IsEnabled = False
            MenuMoveToBottom.IsEnabled = False

            ' MenuSwitch.IsEnabled = False

            Seperator5a.IsEnabled = False

            MenuSort.IsEnabled = False

            ' MenuUndo.IsEnabled = False

        Else

            MenuSort.IsEnabled = True  'v3.4.2

        End If


        MenuExit.IsEnabled = True
        Systray_MenuExit.Enabled = True

        MenuHelp.Foreground = Brushes.Black

        'Match the context menu to the main menu
        MenuContextAdd.IsEnabled = MenuAdd.IsEnabled
        MenuContextEdit.IsEnabled = MenuEdit.IsEnabled
        MenuContextCopy.IsEnabled = MenuCopy.IsEnabled
        MenuContextDelete.IsEnabled = MenuDelete.IsEnabled
        MenuContextMoveUp.IsEnabled = MenuMoveUp.IsEnabled
        MenuContextMoveDown.IsEnabled = MenuMoveDown.IsEnabled
        MenuContextMoveToTop.IsEnabled = MenuMoveToTop.IsEnabled
        MenuContextMoveToBottom.IsEnabled = MenuMoveToBottom.IsEnabled
        MenuContextInsertABlankLine.IsEnabled = MenuInsertABlankLine.IsEnabled
        MenuContextSwitch.IsEnabled = MenuSwitch.IsEnabled
        MenuContextUndo.IsEnabled = MenuUndo.IsEnabled
        MenuContextSwitch.Header = MenuSwitch.Header

    End Sub

    Private Function GetAddChangeInfo(ByVal AddChange As String) As Boolean

        Dim HoldAlwaysOnTop As Boolean = My.Settings.AlwaysOnTop
        Dim HoldWindowSate As WindowState = Me.WindowState

        My.Settings.AlwaysOnTop = True

        gLoadingAddChange = AddChange

        Dim WindowAddChange As WindowAddChange = New WindowAddChange
        gCurrentOwner = WindowAddChange
        WindowAddChange.ShowDialog()

        WindowAddChange = Nothing
        gCurrentOwner = Application.Current.MainWindow

        My.Settings.AlwaysOnTop = HoldAlwaysOnTop

        Me.Visibility = System.Windows.Visibility.Visible
        SeCursor(CursorState.Normal) ' required to allow window to open correctly
        Me.WindowState = HoldWindowSate

        MakeTopMost(SafeNativeMethods.FindWindow(Nothing, Me.Title), My.Settings.AlwaysOnTop)

        Dim ReturnCode As Boolean = (gReturnFromAddChange = "OK")

        Return ReturnCode

    End Function

    Private CurrentSortOrder As Integer = 0


    Public Sub RefreshListView(ByVal SortOrderOfRowToBeSelected As Integer)

        CurrentSortOrder = SortOrderOfRowToBeSelected

        SeCursor(CursorState.Wait)

        LoadListViewFromDatabase()

        For x As Int32 = 0 To ListView1.Items.Count - 1

            Dim aRow As MyTable1ClassForTheListView = ListView1.Items(x)
            If aRow.SortOrder = SortOrderOfRowToBeSelected Then
                ListView1.SelectedIndex = x
                Exit For
            End If

        Next

        ListView1.Focus()

        SeCursor(CursorState.Normal)

    End Sub

#Region "Loading And Saving Column Widths"

    <System.Diagnostics.DebuggerStepThrough()> Private Sub ListView_MouseDoubleClick(ByVal sender As Object, ByVal e As MouseButtonEventArgs) Handles ListView1.MouseDoubleClick

        If (Me.Visibility = System.Windows.Visibility.Hidden) OrElse (Me.WindowState = System.Windows.WindowState.Minimized) Then Exit Sub

        Try
            If TryFindParent(Of GridViewColumnHeader)(TryCast(e.OriginalSource, DependencyObject)) IsNot Nothing Then
                EnforceColumnWidths(TryFindParent(Of GridViewColumnHeader)(TryCast(e.OriginalSource, DependencyObject)))
            End If
        Catch ex As Exception
        End Try

    End Sub

    <System.Diagnostics.DebuggerStepThrough()> Public Shared Function TryFindParent(Of T As Class)(ByVal current As DependencyObject) As T

        Dim parent As DependencyObject = VisualTreeHelper.GetParent(current)

        If parent Is Nothing Then
            Return Nothing
        End If

        If TypeOf parent Is T Then
            Return TryCast(parent, T)
        Else
            Return TryFindParent(Of T)(parent)
        End If

    End Function

    <System.Diagnostics.DebuggerStepThrough()> Private Sub PreventCertainListViewColumnsFromBeingResized(ByVal sender As Object, ByVal e As System.Windows.Controls.Primitives.DragDeltaEventArgs)

        If (Me.Visibility = System.Windows.Visibility.Hidden) OrElse (Me.WindowState = System.Windows.WindowState.Minimized) Then Exit Sub

        Try
            Dim SenderAsThumb As System.Windows.Controls.Primitives.Thumb = TryCast(e.OriginalSource, System.Windows.Controls.Primitives.Thumb)
            EnforceColumnWidths(TryCast(SenderAsThumb.TemplatedParent, GridViewColumnHeader))
        Catch ex As Exception
        End Try

    End Sub

    '<System.Diagnostics.DebuggerStepThrough()>
    Private Sub EnforceColumnWidths(ByVal Header As GridViewColumnHeader)

        ' this only gets triggered when a particular column header is dragged to a new size

        If Header Is Nothing Then Exit Sub

        Try

            Select Case Header.Content.ToString.Trim

                Case Is = "ID", "Sort Order"
                    Header.Column.Width = 0

                Case Is = "Enabled"
                    Header.Column.Width = 75

                Case Is = "Description"
                    If ViewDescription Then
                        If Double.IsNaN(Header.Column.Width) OrElse (Header.Column.Width < 85) Then Header.Column.Width = 85
                    Else
                        Header.Column.Width = 0
                    End If
                    tbFilterDescription.Width = Header.Column.Width

                Case Is = "Listen for"
                    If ViewListenFor Then
                        If Double.IsNaN(Header.Column.Width) OrElse (Header.Column.Width < 85) Then Header.Column.Width = 85
                    Else
                        Header.Column.Width = 0
                    End If
                    tbFilterListenFor.Width = Header.Column.Width

                Case Is = "Open"
                    If ViewOpen Then
                        If Double.IsNaN(Header.Column.Width) OrElse (Header.Column.Width < 85) Then Header.Column.Width = 85
                    Else
                        Header.Column.Width = 0
                    End If
                    tbFilterOpen.Width = Header.Column.Width

                Case Is = "Start directory"
                    If ViewStartIn Then
                        If Double.IsNaN(Header.Column.Width) OrElse (Header.Column.Width < 85) Then Header.Column.Width = 85
                    Else
                        Header.Column.Width = 0
                    End If
                    tbFilterStartIn.Width = Header.Column.Width

                Case Is = "Parameters"
                    If ViewParameters Then
                        If Double.IsNaN(Header.Column.Width) OrElse (Header.Column.Width < 85) Then Header.Column.Width = 85
                    Else
                        Header.Column.Width = 0
                    End If
                    tbFilterParameters.Width = Header.Column.Width

                Case Is = "Admin"
                    If ViewAdmin Then
                        If Double.IsNaN(Header.Column.Width) OrElse (Header.Column.Width < 85) Then Header.Column.Width = 85
                    Else
                        Header.Column.Width = 0
                    End If
                    tbFilterAdmin.Width = Header.Column.Width

                Case Is = "Window state"
                    If ViewStartingWindowState Then
                        If Double.IsNaN(Header.Column.Width) OrElse (Header.Column.Width < 85) Then Header.Column.Width = 85
                    Else
                        Header.Column.Width = 0
                    End If
                    tbFilterStartingWindowState.Width = Header.Column.Width

                Case Is = "Keys to send"
                    If ViewKeysToSend Then
                        If Double.IsNaN(Header.Column.Width) OrElse (Header.Column.Width < 85) Then Header.Column.Width = 85
                    Else
                        Header.Column.Width = 0
                    End If
                    tbFilterKeysToSend.Width = Header.Column.Width

            End Select

        Catch ex As Exception
        End Try

        AdjustFilterPositions()

    End Sub

    Private Sub AdjustFilterPositions()

        DoEvents()

        Const Gap As Double = 10

        Dim Filter1Left, Filter1Right As Double
        Dim Filter2Left, Filter2Right As Double
        Dim Filter3Left, Filter3Right As Double
        Dim Filter4Left, Filter4Right As Double
        Dim Filter5Left, Filter5Right As Double
        Dim Filter6Left, Filter6Right As Double
        Dim Filter7Left, Filter7Right As Double
        Dim Filter8Left, Filter8Right As Double

        Dim Top As Double = tbFilterDescription.Margin.Top
        Dim Bottom As Double = tbFilterDescription.Margin.Bottom

        Dim ColumnHeaders As New List(Of String)
        Dim gv1 As GridView = ListView1.View
        For Each Column In gv1.Columns
            If Column.ActualWidth > 0 Then
                If Column.Header.trim = "Enabled" Then
                Else
                    ColumnHeaders.Add(Column.Header.trim)
                End If
            End If
        Next

        If ColumnHeaders.Count = 0 Then Exit Sub

        Dim gv As GridView = ListView1.View
        If MenuViewDescription.IsChecked Then tbFilterDescription.Width = gv.Columns.Item(ListViewColumns.Description).ActualWidth - Gap
        If MenuViewListenFor.IsChecked Then tbFilterListenFor.Width = gv.Columns.Item(ListViewColumns.ListenFor).ActualWidth - Gap
        If MenuViewOpen.IsChecked Then tbFilterOpen.Width = gv.Columns.Item(ListViewColumns.Open).ActualWidth - Gap
        If MenuViewStartIn.IsChecked Then tbFilterStartIn.Width = gv.Columns.Item(ListViewColumns.StartIn).ActualWidth - Gap
        If MenuViewParameters.IsChecked Then tbFilterParameters.Width = gv.Columns.Item(ListViewColumns.Parameters).ActualWidth - Gap
        If MenuViewAdmin.IsChecked Then tbFilterAdmin.Width = gv.Columns.Item(ListViewColumns.Admin).ActualWidth - Gap
        If MenuViewStartingWindowState.IsChecked Then tbFilterStartingWindowState.Width = gv.Columns.Item(ListViewColumns.StartingWindowState).ActualWidth - Gap
        If MenuViewKeysToSend.IsChecked Then tbFilterKeysToSend.Width = gv.Columns.Item(ListViewColumns.KeysToSend).ActualWidth - Gap

        'there is always at least one column

        'first filter box

        Select Case ColumnHeaders(0)

            Case Is = "Description"
                Filter1Left = tbFilterDescription.Margin.Left
                Filter1Right = Filter1Left + tbFilterDescription.Width - Gap
                tbFilterDescription.Margin = New Thickness(Filter1Left, Top, Filter1Right, Bottom)

            Case Is = "Listen for"
                Filter1Left = tbFilterListenFor.Margin.Left
                Filter1Right = Filter1Left + tbFilterListenFor.Width - Gap
                tbFilterListenFor.Margin = New Thickness(Filter1Left, Top, Filter1Right, Bottom)

            Case Is = "Open"
                Filter1Left = tbFilterOpen.Margin.Left
                Filter1Right = Filter1Left + tbFilterOpen.Width - Gap
                tbFilterOpen.Margin = New Thickness(Filter1Left, Top, Filter1Right, Bottom)

            Case Is = "Start directory"
                Filter1Left = tbFilterStartIn.Margin.Left
                Filter1Right = Filter1Left + tbFilterStartIn.Width - Gap
                tbFilterStartIn.Margin = New Thickness(Filter1Left, Top, Filter1Right, Bottom)

            Case Is = "Parameters"
                Filter1Left = tbFilterParameters.Margin.Left
                Filter1Right = Filter1Left + tbFilterParameters.Width - Gap
                tbFilterParameters.Margin = New Thickness(Filter1Left, Top, Filter1Right, Bottom)

            Case Is = "Admin"
                Filter1Left = tbFilterAdmin.Margin.Left
                Filter1Right = Filter1Left + tbFilterAdmin.Width - Gap
                tbFilterAdmin.Margin = New Thickness(Filter1Left, Top, Filter1Right, Bottom)

            Case Is = "Window state"
                Filter1Left = tbFilterStartingWindowState.Margin.Left
                Filter1Right = Filter1Left + tbFilterStartingWindowState.Width - Gap
                tbFilterStartingWindowState.Margin = New Thickness(Filter1Left, Top, Filter1Right, Bottom)

            Case Is = "Keys to send"
                Filter1Left = tbFilterStartingWindowState.Margin.Left
                Filter1Right = Filter1Left + tbFilterKeysToSend.Width - Gap
                tbFilterKeysToSend.Margin = New Thickness(Filter1Left, Top, Filter1Right, Bottom)

        End Select

        'second filter box

        If ColumnHeaders.Count < 2 Then
            Exit Sub

        Else

            Select Case ColumnHeaders(1)

                Case Is = "Description"

                Case Is = "Listen for"
                    Filter2Left = Filter1Right + 2 * Gap
                    Filter2Right = Filter2Left + tbFilterListenFor.Width - Gap
                    tbFilterListenFor.Margin = New Thickness(Filter2Left, Top, Filter2Right, Bottom)

                Case Is = "Open"
                    Filter2Left = Filter1Right + 2 * Gap
                    Filter2Right = Filter2Left + tbFilterOpen.Width - Gap
                    tbFilterOpen.Margin = New Thickness(Filter2Left, Top, Filter2Right, Bottom)

                Case Is = "Start directory"
                    Filter2Left = Filter1Right + 2 * Gap
                    Filter2Right = Filter2Left + tbFilterStartIn.Width - Gap
                    tbFilterStartIn.Margin = New Thickness(Filter2Left, Top, Filter2Right, Bottom)

                Case Is = "Parameters"
                    Filter2Left = Filter1Right + 2 * Gap
                    Filter2Right = Filter2Left + tbFilterParameters.Width - Gap
                    tbFilterParameters.Margin = New Thickness(Filter2Left, Top, Filter2Right, Bottom)

                Case Is = "Admin"
                    Filter2Left = Filter1Right + 2 * Gap
                    Filter2Right = Filter2Left + tbFilterAdmin.Width - Gap
                    tbFilterAdmin.Margin = New Thickness(Filter2Left, Top, Filter2Right, Bottom)

                Case Is = "Window state"
                    Filter2Left = Filter1Right + 2 * Gap
                    Filter2Right = Filter2Left + tbFilterStartingWindowState.Width - Gap
                    tbFilterStartingWindowState.Margin = New Thickness(Filter2Left, Top, Filter2Right, Bottom)

                Case Is = "Keys to send"
                    Filter2Left = Filter1Right + 2 * Gap
                    Filter2Right = Filter2Left + tbFilterKeysToSend.Width - Gap
                    tbFilterKeysToSend.Margin = New Thickness(Filter2Left, Top, Filter2Right, Bottom)

            End Select

        End If

        'third filter box

        If ColumnHeaders.Count < 3 Then
            Exit Sub

        Else

            Select Case ColumnHeaders(2)

                Case Is = "Open"
                    Filter3Left = Filter2Right + 2 * Gap
                    Filter3Right = Filter3Left + tbFilterOpen.Width - Gap
                    tbFilterOpen.Margin = New Thickness(Filter3Left, Top, Filter3Right, Bottom)

                Case Is = "Start directory"
                    Filter3Left = Filter2Right + 2 * Gap
                    Filter3Right = Filter3Left + tbFilterStartIn.Width - Gap
                    tbFilterStartIn.Margin = New Thickness(Filter3Left, Top, Filter3Right, Bottom)

                Case Is = "Parameters"
                    Filter3Left = Filter2Right + 2 * Gap
                    Filter3Right = Filter3Left + tbFilterParameters.Width - Gap
                    tbFilterParameters.Margin = New Thickness(Filter3Left, Top, Filter3Right, Bottom)

                Case Is = "Admin"
                    Filter3Left = Filter2Right + 2 * Gap
                    Filter3Right = Filter3Left + tbFilterAdmin.Width - Gap
                    tbFilterAdmin.Margin = New Thickness(Filter3Left, Top, Filter3Right, Bottom)

                Case Is = "Window state"
                    Filter3Left = Filter2Right + 2 * Gap
                    Filter3Right = Filter3Left + tbFilterStartingWindowState.Width - Gap
                    tbFilterStartingWindowState.Margin = New Thickness(Filter3Left, Top, Filter3Right, Bottom)

                Case Is = "Keys to send"
                    Filter3Left = Filter2Right + 2 * Gap
                    Filter3Right = Filter3Left + tbFilterKeysToSend.Width - Gap
                    tbFilterKeysToSend.Margin = New Thickness(Filter3Left, Top, Filter3Right, Bottom)

            End Select

        End If

        'forth filter box

        If ColumnHeaders.Count < 4 Then

            Exit Sub

        Else


            Select Case ColumnHeaders(3)

                Case Is = "Start directory"
                    Filter4Left = Filter3Right + 2 * Gap
                    Filter4Right = Filter4Left + tbFilterStartIn.Width - Gap
                    tbFilterStartIn.Margin = New Thickness(Filter4Left, Top, Filter4Right, Bottom)

                Case Is = "Parameters"
                    Filter4Left = Filter3Right + 2 * Gap
                    Filter4Right = Filter4Left + tbFilterParameters.Width - Gap
                    tbFilterParameters.Margin = New Thickness(Filter4Left, Top, Filter4Right, Bottom)

                Case Is = "Admin"
                    Filter4Left = Filter3Right + 2 * Gap
                    Filter4Right = Filter4Left + tbFilterAdmin.Width - Gap
                    tbFilterAdmin.Margin = New Thickness(Filter4Left, Top, Filter4Right, Bottom)

                Case Is = "Window state"
                    Filter4Left = Filter3Right + 2 * Gap
                    Filter4Right = Filter4Left + tbFilterStartingWindowState.Width - Gap
                    tbFilterStartingWindowState.Margin = New Thickness(Filter4Left, Top, Filter4Right, Bottom)

                Case Is = "Keys to send"
                    Filter4Left = Filter3Right + 2 * Gap
                    Filter4Right = Filter4Left + tbFilterKeysToSend.Width - Gap
                    tbFilterKeysToSend.Margin = New Thickness(Filter4Left, Top, Filter4Right, Bottom)

            End Select

        End If

        'fifth filter box

        If ColumnHeaders.Count < 5 Then

            Exit Sub

        Else

            Select Case ColumnHeaders(4)

                Case Is = "Parameters"
                    Filter5Left = Filter4Right + 2 * Gap
                    Filter5Right = Filter5Left + tbFilterParameters.Width - Gap
                    tbFilterParameters.Margin = New Thickness(Filter5Left, Top, Filter5Right, Bottom)

                Case Is = "Admin"
                    Filter5Left = Filter4Right + 2 * Gap
                    Filter5Right = Filter5Left + tbFilterAdmin.Width - Gap
                    tbFilterAdmin.Margin = New Thickness(Filter5Left, Top, Filter5Right, Bottom)

                Case Is = "Window state"
                    Filter5Left = Filter4Right + 2 * Gap
                    Filter5Right = Filter5Left + tbFilterStartingWindowState.Width - Gap
                    tbFilterStartingWindowState.Margin = New Thickness(Filter5Left, Top, Filter5Right, Bottom)

                Case Is = "Keys to send"
                    Filter5Left = Filter4Right + 2 * Gap
                    Filter5Right = Filter5Left + tbFilterKeysToSend.Width - Gap
                    tbFilterKeysToSend.Margin = New Thickness(Filter5Left, Top, Filter5Right, Bottom)

            End Select

        End If

        'sixth filter box

        If ColumnHeaders.Count < 6 Then

            Exit Sub

        Else

            Select Case ColumnHeaders(5)

                Case Is = "Admin"
                    Filter6Left = Filter5Right + 2 * Gap
                    Filter6Right = Filter6Left + tbFilterAdmin.Width - Gap
                    tbFilterAdmin.Margin = New Thickness(Filter6Left, Top, Filter6Right, Bottom)

                Case Is = "Window state"
                    Filter6Left = Filter5Right + 2 * Gap
                    Filter6Right = Filter6Left + tbFilterStartingWindowState.Width - Gap
                    tbFilterStartingWindowState.Margin = New Thickness(Filter6Left, Top, Filter6Right, Bottom)

                Case Is = "Keys to send"
                    Filter6Left = Filter5Right + 2 * Gap
                    Filter6Right = Filter6Left + tbFilterKeysToSend.Width - Gap
                    tbFilterKeysToSend.Margin = New Thickness(Filter6Left, Top, Filter6Right, Bottom)

            End Select

        End If

        'seventh filter box

        If ColumnHeaders.Count < 7 Then

            Exit Sub

        Else

            Select Case ColumnHeaders(6)

                Case Is = "Window state"
                    Filter7Left = Filter6Right + 2 * Gap
                    Filter7Right = Filter7Left + tbFilterStartingWindowState.Width - Gap
                    tbFilterStartingWindowState.Margin = New Thickness(Filter7Left, Top, Filter7Right, Bottom)

                Case Is = "Keys to send"
                    Filter7Left = Filter6Right + 2 * Gap
                    Filter7Right = Filter7Left + tbFilterKeysToSend.Width - Gap
                    tbFilterKeysToSend.Margin = New Thickness(Filter7Left, Top, Filter7Right, Bottom)

            End Select

        End If

        'eight filter box

        If ColumnHeaders.Count < 8 Then

            Exit Sub

        Else

            Select Case ColumnHeaders(7)

                Case Is = "Keys to send"
                    Filter8Left = Filter7Right + 2 * Gap
                    Filter8Right = Filter8Left + tbFilterKeysToSend.Width - Gap
                    tbFilterKeysToSend.Margin = New Thickness(Filter8Left, Top, Filter8Right, Bottom)

            End Select

        End If

    End Sub

    Private Sub LoadWindowLocationSizeAndColumnWidths()

        If (Me.Visibility = System.Windows.Visibility.Hidden) OrElse (Me.WindowState = System.Windows.WindowState.Minimized) Then Exit Sub

        On Error Resume Next

        'Set starting window location

        Dim dwa = System.Windows.SystemParameters.WorkArea

        Dim WillFitInWindow As Boolean = False

        If My.Settings.Top >= dwa.Top Then
            If My.Settings.Top <= (dwa.Top + dwa.Height - 40) Then
                If My.Settings.Left >= dwa.Left Then
                    If My.Settings.Left <= (dwa.Left + dwa.Width - 40) Then
                        WillFitInWindow = True
                    End If
                End If
            End If
        End If

        If WillFitInWindow Then
            Me.Top = My.Settings.Top
            Me.Left = My.Settings.Left
        Else
            Me.Top = 50
            Me.Left = 50
            My.Settings.Top = 50
            My.Settings.Left = 50
        End If

        'Set Window Size

        If My.Settings.MainWindowSize.Width < Me.MinWidth Then
            Me.Width = Me.MinWidth
        Else
            Me.Width = My.Settings.MainWindowSize.Width
        End If

        If My.Settings.MainWindowSize.Height < Me.MinHeight Then
            Me.Height = Me.MinHeight
        Else
            Me.Height = My.Settings.MainWindowSize.Height
        End If

        'Set rest of window
        MenuViewDescription.IsChecked = My.Settings.ViewDescription
        MenuViewListenFor.IsChecked = My.Settings.ViewListenFor
        MenuViewOpen.IsChecked = My.Settings.ViewOpen
        MenuViewStartIn.IsChecked = My.Settings.ViewStartIn
        MenuViewParameters.IsChecked = My.Settings.ViewParameters
        MenuViewAdmin.IsChecked = My.Settings.ViewAdmin
        MenuViewStartingWindowState.IsChecked = My.Settings.ViewStartingWindowState
        MenuViewKeysToSend.IsChecked = My.Settings.ViewKeysToSend
        MenuViewDisabledCards.IsChecked = My.Settings.ViewDisabledCards
        MenuViewFilters.IsChecked = My.Settings.ViewFilters

        ViewDescription = MenuViewDescription.IsChecked
        ViewListenFor = MenuViewListenFor.IsChecked
        ViewOpen = MenuViewOpen.IsChecked
        ViewStartIn = MenuViewStartIn.IsChecked
        ViewParameters = MenuViewParameters.IsChecked
        ViewAdmin = MenuViewAdmin.IsChecked
        ViewStartingWindowState = MenuViewStartingWindowState.IsChecked
        ViewKeysToSend = MenuViewKeysToSend.IsChecked

        If My.Settings.ViewDescription OrElse My.Settings.ViewListenFor OrElse My.Settings.ViewOpen OrElse My.Settings.ViewParameters OrElse My.Settings.ViewStartIn OrElse My.Settings.ViewAdmin OrElse
            My.Settings.ViewStartingWindowState OrElse My.Settings.ViewKeysToSend Then
        Else
            MenuViewListenFor.IsChecked = True 'this should never happened, defensive coding
        End If

        Dim gv As GridView = ListView1.View

        ' Width is set to -1 in the initial load, and is used to signify the column should be autosized

        SetColumnWidths_Loading(gv, My.Settings.ViewDescription, ListViewColumns.Description, My.Settings.ColumnWidthDescription)
        SetColumnWidths_Loading(gv, My.Settings.ViewListenFor, ListViewColumns.ListenFor, My.Settings.ColumnWidthListenFor)
        SetColumnWidths_Loading(gv, My.Settings.ViewOpen, ListViewColumns.Open, My.Settings.ColumnWidthOpen)
        SetColumnWidths_Loading(gv, My.Settings.ViewStartIn, ListViewColumns.StartIn, My.Settings.ColumnWidthStartIn)
        SetColumnWidths_Loading(gv, My.Settings.ViewParameters, ListViewColumns.Parameters, My.Settings.ColumnWidthParameters)
        SetColumnWidths_Loading(gv, My.Settings.ViewAdmin, ListViewColumns.Admin, My.Settings.ColumnWidthAdmin)
        SetColumnWidths_Loading(gv, My.Settings.ViewStartingWindowState, ListViewColumns.StartingWindowState, My.Settings.ColumnWidthStartingWindowState)
        SetColumnWidths_Loading(gv, My.Settings.ViewKeysToSend, ListViewColumns.KeysToSend, My.Settings.ColumnWidthKeysToSend)

        ListView1.View = gv

        gv = Nothing

        If MenuViewFilters.IsChecked Then
        Else
            ShowOrHideFilters()
        End If

        AdjustFilterPositions()
        SetFiltersVisibility()

        Static OnlyDoOnce As Boolean = True

        If OnlyDoOnce Then
            Me.ListView1.AddHandler(System.Windows.Controls.Primitives.Thumb.DragDeltaEvent, New System.Windows.Controls.Primitives.DragDeltaEventHandler(AddressOf PreventCertainListViewColumnsFromBeingResized), True)
            OnlyDoOnce = False
        End If

    End Sub

    Private Sub SetColumnWidths_Loading(ByRef gv As GridView, ByVal ViewParm As Boolean, ByVal lvColumnNumber As Integer, ByVal ColumnnWidth As Double)


        If ViewParm Then

            If ColumnnWidth > 0 Then
                gv.Columns.Item(lvColumnNumber).Width = ColumnnWidth
            Else
                gv.Columns.Item(lvColumnNumber).Width = System.Double.NaN
            End If

        Else

            If (gv.Columns(lvColumnNumber).Header.Trim = "Description") AndAlso (lvColumnNumber = ListViewColumns.Description) Then
                gv.Columns.Item(lvColumnNumber).Width = 0

            ElseIf (gv.Columns(lvColumnNumber).Header.Trim = "Listen for") AndAlso (lvColumnNumber = ListViewColumns.ListenFor) Then
                gv.Columns.Item(lvColumnNumber).Width = 0

            ElseIf (gv.Columns(lvColumnNumber).Header.Trim = "Open") AndAlso (lvColumnNumber = ListViewColumns.Open) Then
                gv.Columns.Item(lvColumnNumber).Width = 0

            ElseIf (gv.Columns(lvColumnNumber).Header.Trim = "Start directory") AndAlso (lvColumnNumber = ListViewColumns.StartIn) Then
                gv.Columns.Item(lvColumnNumber).Width = 0

            ElseIf (gv.Columns(lvColumnNumber).Header.Trim = "Parameters") AndAlso (lvColumnNumber = ListViewColumns.Parameters) Then
                gv.Columns.Item(lvColumnNumber).Width = 0

            ElseIf (gv.Columns(lvColumnNumber).Header.Trim = "Admin") AndAlso (lvColumnNumber = ListViewColumns.Admin) Then
                gv.Columns.Item(lvColumnNumber).Width = 0

            ElseIf (gv.Columns(lvColumnNumber).Header.Trim = "Window state") AndAlso (lvColumnNumber = ListViewColumns.StartingWindowState) Then
                gv.Columns.Item(lvColumnNumber).Width = 0

            ElseIf (gv.Columns(lvColumnNumber).Header.Trim = "Keys to send") AndAlso (lvColumnNumber = ListViewColumns.KeysToSend) Then
                gv.Columns.Item(lvColumnNumber).Width = 0

            End If

        End If

        SetFiltersVisibility()

    End Sub

    Private Sub SetFiltersVisibility()

        If MenuViewDescription.IsChecked Then
            tbFilterDescription.Visibility = Visibility.Visible
        Else
            tbFilterDescription.Visibility = Visibility.Hidden
        End If

        If MenuViewListenFor.IsChecked Then
            tbFilterListenFor.Visibility = Visibility.Visible
        Else
            tbFilterListenFor.Visibility = Visibility.Hidden
        End If

        If MenuViewOpen.IsChecked Then
            tbFilterOpen.Visibility = Visibility.Visible
        Else
            tbFilterOpen.Visibility = Visibility.Hidden
        End If

        If MenuViewStartIn.IsChecked Then
            tbFilterStartIn.Visibility = Visibility.Visible
        Else
            tbFilterStartIn.Visibility = Visibility.Hidden
        End If

        If MenuViewParameters.IsChecked Then
            tbFilterParameters.Visibility = Visibility.Visible
        Else
            tbFilterParameters.Visibility = Visibility.Hidden
        End If

        If MenuViewAdmin.IsChecked Then
            tbFilterAdmin.Visibility = Visibility.Visible
        Else
            tbFilterAdmin.Visibility = Visibility.Hidden
        End If

        If MenuViewStartingWindowState.IsChecked Then
            tbFilterStartingWindowState.Visibility = Visibility.Visible
        Else
            tbFilterStartingWindowState.Visibility = Visibility.Hidden
        End If

        If MenuViewKeysToSend.IsChecked Then
            tbFilterKeysToSend.Visibility = Visibility.Visible
        Else
            tbFilterKeysToSend.Visibility = Visibility.Hidden
        End If

        AdjustFilterPositions()

    End Sub

    Private Sub SaveWindowLocationSizeAndColumnWidthsAndOthers()

        If (Me.Visibility = System.Windows.Visibility.Hidden) OrElse (Me.Visibility = System.Windows.Visibility.Collapsed) OrElse (Me.WindowState = System.Windows.WindowState.Minimized) Then Exit Sub

        Try

            Dim SaveRequired As Boolean = False 'v 3.5.3 only save settings if a change is needed

            If My.Settings.Top = Me.Top Then
            Else
                My.Settings.Top = Me.Top
                SaveRequired = True
            End If

            If My.Settings.Left = Me.Left Then
            Else
                My.Settings.Left = Me.Left
                SaveRequired = True
            End If

            Try
                If Me.WindowState <> WindowState.Maximized Then
                    If (Me.Width > 0) AndAlso (Me.Height > 0) Then
                        If My.Settings.MainWindowSize = New System.Drawing.Size(Me.Width, Me.Height) Then
                        Else
                            My.Settings.MainWindowSize = New System.Drawing.Size(Me.Width, Me.Height)
                            SaveRequired = True
                        End If
                    End If
                End If
            Catch ex As Exception
            End Try

            If My.Settings.ViewDescription = MenuViewDescription.IsChecked Then
            Else
                My.Settings.ViewDescription = MenuViewDescription.IsChecked
                SaveRequired = True
            End If

            If My.Settings.ViewListenFor = MenuViewListenFor.IsChecked Then
            Else
                My.Settings.ViewListenFor = MenuViewListenFor.IsChecked
                SaveRequired = True
            End If

            If My.Settings.ViewOpen = MenuViewOpen.IsChecked Then
            Else
                My.Settings.ViewOpen = MenuViewOpen.IsChecked
                SaveRequired = True
            End If

            If My.Settings.ViewStartIn = MenuViewStartIn.IsChecked Then
            Else
                My.Settings.ViewStartIn = MenuViewStartIn.IsChecked
                SaveRequired = True
            End If

            If My.Settings.ViewParameters = MenuViewParameters.IsChecked Then
            Else
                My.Settings.ViewParameters = MenuViewParameters.IsChecked
                SaveRequired = True
            End If

            If My.Settings.ViewAdmin = MenuViewAdmin.IsChecked Then
            Else
                My.Settings.ViewAdmin = MenuViewAdmin.IsChecked
                SaveRequired = True
            End If

            If My.Settings.ViewStartingWindowState = MenuViewStartingWindowState.IsChecked Then
            Else
                My.Settings.ViewStartingWindowState = MenuViewStartingWindowState.IsChecked
                SaveRequired = True
            End If

            If My.Settings.ViewKeysToSend = MenuViewKeysToSend.IsChecked Then
            Else
                My.Settings.ViewKeysToSend = MenuViewKeysToSend.IsChecked
                SaveRequired = True
            End If


            If My.Settings.ViewDisabledCards = MenuViewDisabledCards.IsChecked Then
            Else
                My.Settings.ViewDisabledCards = MenuViewDisabledCards.IsChecked
                SaveRequired = True
            End If

            If My.Settings.ViewFilters = MenuViewFilters.IsChecked Then
            Else
                My.Settings.ViewFilters = MenuViewFilters.IsChecked
                SaveRequired = True
            End If

            Dim gv As GridView = ListView1.View
            Dim WorkingWidth As Double

            'v1.8

            If MenuViewDescription.IsChecked Then
                SetColumnWidths_Saving(gv, ListViewColumns.Description, WorkingWidth)
                If My.Settings.ColumnWidthDescription = WorkingWidth Then
                Else
                    My.Settings.ColumnWidthDescription = WorkingWidth
                    SaveRequired = True
                End If
            End If

            If MenuViewListenFor.IsChecked Then

                SetColumnWidths_Saving(gv, ListViewColumns.ListenFor, WorkingWidth)
                If My.Settings.ColumnWidthListenFor = WorkingWidth Then
                Else
                    My.Settings.ColumnWidthListenFor = WorkingWidth
                    SaveRequired = True
                End If

            End If

            If MenuViewOpen.IsChecked Then
                SetColumnWidths_Saving(gv, ListViewColumns.Open, WorkingWidth)
                If My.Settings.ColumnWidthOpen = WorkingWidth Then
                Else
                    My.Settings.ColumnWidthOpen = WorkingWidth
                    SaveRequired = True
                End If
            End If

            If MenuViewStartIn.IsChecked Then
                SetColumnWidths_Saving(gv, ListViewColumns.StartIn, WorkingWidth)
                If My.Settings.ColumnWidthStartIn = WorkingWidth Then
                Else
                    My.Settings.ColumnWidthStartIn = WorkingWidth
                    SaveRequired = True
                End If
            End If

            If MenuViewParameters.IsChecked Then
                SetColumnWidths_Saving(gv, ListViewColumns.Parameters, WorkingWidth)
                If My.Settings.ColumnWidthParameters = WorkingWidth Then
                Else
                    My.Settings.ColumnWidthParameters = WorkingWidth
                    SaveRequired = True
                End If
            End If

            If MenuViewAdmin.IsChecked Then
                SetColumnWidths_Saving(gv, ListViewColumns.Admin, WorkingWidth)
                If My.Settings.ColumnWidthAdmin = WorkingWidth Then
                Else
                    My.Settings.ColumnWidthAdmin = WorkingWidth
                    SaveRequired = True
                End If
            End If

            If MenuViewStartingWindowState.IsChecked Then
                SetColumnWidths_Saving(gv, ListViewColumns.StartingWindowState, WorkingWidth)
                If My.Settings.ColumnWidthStartingWindowState = WorkingWidth Then
                Else
                    My.Settings.ColumnWidthStartingWindowState = WorkingWidth
                    SaveRequired = True
                End If
            End If

            If MenuViewKeysToSend.IsChecked Then
                SetColumnWidths_Saving(gv, ListViewColumns.KeysToSend, WorkingWidth)
                If My.Settings.ColumnWidthKeysToSend = WorkingWidth Then
                Else
                    My.Settings.ColumnWidthKeysToSend = WorkingWidth
                    SaveRequired = True
                End If
            End If

            If My.Settings.SortByDescription = MenuSort.IsChecked Then
            Else
                My.Settings.SortByDescription = MenuSort.IsChecked
                SaveRequired = True
            End If

            gMenuSort = MenuSort.IsChecked 'v2.5.3

            If SaveRequired Then 'v3.5.3 only save settings if a change is needed
                My.Settings.Save()
            End If

            gv = Nothing

        Catch ex As Exception

            Log("Problem saving settings" & vbCrLf & ex.ToString)

        End Try

    End Sub

    Private Sub SetColumnWidths_Saving(ByVal gv As GridView, ByVal lvColumnNumber As Integer, ByRef ColumnnWidth As Double)

        Try
            If Double.IsNaN(gv.Columns.Item(lvColumnNumber).Width) Then
                ColumnnWidth = -1
            Else
                ColumnnWidth = gv.Columns.Item(lvColumnNumber).Width
            End If
        Catch ex As Exception
            ColumnnWidth = -1
        End Try

    End Sub

#End Region

#Region "Undo Logic"

    Enum UndoRational
        add = 1
        delete = 2
        copy = 3
        insert_a_blank_line = 4
        move_up = 5
        move_down = 6
        move_top = 7
        move_bottom = 8
        sort = 9
        remove_sort = 10
        toggle_Master_Switch = 11
        toggle_a_switch = 12
        change_to_existing_Push2Run_card = 13
        drop = 14
        import = 15
    End Enum

    Public Structure UndoTableStructure
        Public Rational As UndoRational
        Public UndoSortOrderOfSelectedRecord As Integer
        Friend TableEntries() As gRowRecord
    End Structure

    Private MyUndoTableIndex As Int32 = 0

    Private MyUndoTableCurrentDimension As Int32 = 100
    Private MyUndoTable(MyUndoTableCurrentDimension) As UndoTableStructure

    Private Const MyUndoTableDimensionIncrement As Int32 = 100


    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub AddToUndoTable(ByVal Rational As UndoRational)

        Try

            ' add current listview to table
            MyUndoTableIndex += 1

            ' ensure the undo table will not overflow
            If MyUndoTableIndex = MyUndoTableCurrentDimension - 1 Then
                MyUndoTableCurrentDimension += MyUndoTableDimensionIncrement
                ReDim Preserve MyUndoTable(MyUndoTableCurrentDimension)
            End If

            ' update MyUndoTable

            MyUndoTable(MyUndoTableIndex).Rational = Rational
            MyUndoTable(MyUndoTableIndex).UndoSortOrderOfSelectedRecord = gCurrentlySelectedRow.SortOrder

            If MenuViewFilters.IsChecked Then
                LoadUndoTableFromDatabase()
            Else
                LoadUndoTableFromListView()
            End If

        Catch ex As Exception

        End Try

    End Sub


    Private Sub LoadUndoTableFromDatabase()

        Try

            Dim DatabaseSource As New List(Of MyTable1ClassForTheListView)
            DatabaseSource = LoadDatabaseIntoAList(True)

            ReDim MyUndoTable(MyUndoTableIndex).TableEntries(DatabaseSource.Count - 1)

            Dim x As Integer = 0
            For Each Entry In DatabaseSource

                With MyUndoTable(MyUndoTableIndex).TableEntries(x)

                    .ID = Entry.ID
                    .SortOrder = Entry.SortOrder
                    .DesiredStatus = Entry.DesiredStatus
                    .WorkingStatus = Entry.WorkingStatus
                    .Description = Entry.Description
                    .ListenFor = Entry.ListenFor
                    .Open = Entry.Open
                    .Parameters = Entry.Parameters
                    .StartIn = Entry.StartIn
                    .Admin = Entry.Admin
                    .StartingWindowState = Entry.StartingWindowState
                    .KeysToSend = Entry.KeysToSend

                End With
                x += 1

            Next

            DatabaseSource = Nothing

        Catch ex As Exception
            Dim Result As MessageBoxResult = TopMostMessageBox(gCurrentOwner, ex.Message.ToString, "Push2Run - Error", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK, System.Windows.MessageBoxOptions.None)
        End Try

    End Sub

    Private Sub LoadUndoTableFromListView()

        Try

            ReDim MyUndoTable(MyUndoTableIndex).TableEntries(ListView1.Items.Count - 1)

            Dim x As Integer = 0
            For Each Entry In ListView1.Items

                With MyUndoTable(MyUndoTableIndex).TableEntries(x)

                    .ID = Entry.ID
                    .SortOrder = Entry.SortOrder
                    .DesiredStatus = Entry.DesiredStatus
                    .WorkingStatus = Entry.WorkingStatus
                    .Description = Entry.Description
                    .ListenFor = Entry.ListenFor
                    .Open = Entry.Open
                    .Parameters = Entry.Parameters
                    .StartIn = Entry.StartIn
                    .Admin = Entry.Admin
                    .StartingWindowState = Entry.StartingWindowState
                    .KeysToSend = Entry.KeysToSend

                End With
                x += 1

            Next

        Catch ex As Exception

            Dim Result As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "Error adding to undo table from list view" & vbCrLf & ex.Message.ToString, "Push2Run - Error", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK, System.Windows.MessageBoxOptions.None)

        End Try

    End Sub


    Private Sub CancelAddToUndoTable()

        If MyUndoTableIndex > 0 Then
            MyUndoTableIndex -= 1
        End If

    End Sub

    Private Function RestoreFromUndoTable() As Integer

        Dim ReturnValue As Integer = 0

        Try

            'Step 1: Create a backup of the database 
            If File.Exists(gSQLiteFullDatabaseName & ".backup") Then File.Delete(gSQLiteFullDatabaseName & ".backup")
            File.Copy(gSQLiteFullDatabaseName, gSQLiteFullDatabaseName & ".backup", True)

            'Step 2: Drop Table1 
            RunSQL("DROP TABLE Table1 ;", False)

            'Step 3: Recreate Table1
            CreateTable1()

            'Step 4: Reload Table1
            InsertManyRecords(MyUndoTable(MyUndoTableIndex).TableEntries)

            'Step 5: set current sortorder value so listview can be repositioned 
            ReturnValue = MyUndoTable(MyUndoTableIndex).UndoSortOrderOfSelectedRecord

            'Step 6: back down the Undo table pointer
            MyUndoTableIndex -= 1

        Catch ex As Exception

            'Restore from backup
            If File.Exists(gSQLiteFullDatabaseName & ".backup") Then

                If File.Exists(gSQLiteFullDatabaseName) Then
                    'Restore from backup
                    File.Delete(gSQLiteFullDatabaseName)
                    File.Copy(gSQLiteFullDatabaseName & ".backup", gSQLiteFullDatabaseName)

                End If

            End If

            MyUndoTableIndex = 0

        Finally

            If File.Exists(gSQLiteFullDatabaseName & ".backup") Then
                File.Delete(gSQLiteFullDatabaseName & ".backup")
            End If

        End Try

        'Return the Sort Order of the Row to be selected

        Return ReturnValue

    End Function

#End Region

#Region "Move form"

    Private Sub Me_MouseLeftButtonDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles Me.MouseLeftButtonDown

        Try
            DragMove()
        Catch ex As Exception
        End Try

    End Sub

#End Region

    Private gPasswordConfirmationFileName As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData) & "\Push2Run\Push2Run.dat"

    Private WithEvents WorkerTimer As System.Windows.Forms.Timer = New System.Windows.Forms.Timer

    Private WithEvents KeepAccountActiveTimer As System.Windows.Forms.Timer = New System.Windows.Forms.Timer

    Private Sub KeepAccountActiveTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KeepAccountActiveTimer.Tick

        RequestPushbulletToKeepAccountActive()

    End Sub

    Private objlock4 As Object = New Object
    Private Sub RequestPushbulletToKeepAccountActive()

        SyncLock (objlock4)

            Try

                Dim TwelveHoursAgo = Now.AddMilliseconds(-1 * gTwelveHoursInMilliSeconds)

                If My.Settings.LastRequesttoKeepPushbulletToKeepAccountActive >= TwelveHoursAgo Then
                    Exit Try
                End If

                My.Settings.LastRequesttoKeepPushbulletToKeepAccountActive = Now
                My.Settings.Save()

                KeepAccountActiveTimer.Interval = gTwelveHoursInMilliSeconds

                If PushbulletEnbledInTesting Then

                    If KeepPushbulletAccountActive(EncryptionClass.Decrypt(My.Settings.PushBulletAPI)) Then
                        Log("Requested Pushbullet keep your account active")
                        Log("")
                    Else
                        Log("Request to Pushbullet to keep your account active failed")
                        Log("")
                    End If

                End If

            Catch ex As Exception

            End Try

        End SyncLock

    End Sub

    Private MaxMemoryAllocation As Int32 = 125000 'made large for rich text fields
    Private TextBufferHandle As IntPtr = Marshal.AllocHGlobal(MaxMemoryAllocation)


    Private Sub SetPriority(ByVal Prioirty As ProcessPriorityClass)

        Static Dim myProcess As Process = Process.GetCurrentProcess
        myProcess.PriorityClass = Prioirty

    End Sub

#End Region

#Region "Pushbullet Related"

    Private Sub Websocket_Opened_Pushbullet(ByVal sender As Object, ByVal e As EventArgs)
        'Log("")
    End Sub

    Private Sub Websocket_MessageReceived_Pushbullet(ByVal sender As WebSocket4Net.WebSocket, ByVal e As WebSocket4Net.MessageReceivedEventArgs)

        LastTimeDataWasReceivedFromPushbullet = Now

        Static Dim LastIden As String = String.Empty

        'v2.2 change to test for the string without spaces 
        Dim MessageIn As String = e.Message.ToString
        MessageIn = RegularExpressions.Regex.Replace(MessageIn, "\s+", "")  'remove all spaces

        Select Case MessageIn 'e.Message.ToString

            'Case = "{""type"":    ""nop""}"

            Case = "{""type"":""tickle"",""subtype"":""push""}" '"{""type"" ""tickle"", ""subtype"": ""push""}", "{""type"":""tickle"",""subtype"":""push""}", "{""type"": ""tickle"", ""subtype"": ""push""}"

                Dim MostRecentNote As NoteInfo = GetNotesSinceLastRequest()

                If MostRecentNote.Iden = LastIden Then
                    Exit Select 'ignore duplicate pushes 'v2.0.5
                Else
                    LastIden = MostRecentNote.Iden
                End If

                If MostRecentNote.Dismissed Then Exit Select 'ignore dismissed pushes

                Log("")
                Log("Incoming Pushbullet push ...")
                Log("<Title>" & vbTab & MostRecentNote.Title)
                Log("<Body>" & vbTab & MostRecentNote.Body)

                If (MostRecentNote.Title Is Nothing) OrElse (MostRecentNote.Title = String.Empty) Then
                    Log("Title is blank - no further action will be taken")
                    Log("")
                    Exit Select
                End If

                If MatchOnTitleFilter(MostRecentNote.Title.ToUpper) Then
                Else
                    Log("Title did not match Push2Run's title filter - no further action will be taken")
                    Log("")
                    Exit Select
                End If

                'Automatically dismiss the pop-up notification
                DissmissANote(EncryptionClass.Decrypt(My.Settings.PushBulletAPI), MostRecentNote.Iden)

                If (MostRecentNote.Body Is Nothing) OrElse (MostRecentNote.Body = String.Empty) Then
                    Log("Body was empty - no further action will be taken")
                    Log("")
                    Exit Select
                End If

                If My.Settings.UsePushbullet Then
                    ActionIncomingMessage(MessageSource.Pushbullet, MostRecentNote.Body)
                Else
                    Log("Pushbullet not enabled in options - no further action will be taken")
                End If

        End Select

    End Sub


    Private Sub Websocket_Opened_Pushover(ByVal sender As Object, ByVal e As EventArgs)

    End Sub

    Private Sub Websocket_MessageReceived_Pushover(ByVal sender As WebSocket4Net.WebSocket, ByVal e As WebSocket4Net.MessageReceivedEventArgs)

    End Sub

    Private Function MatchOnTitleFilter(ByVal TitleIn As String) As Boolean

        Dim ReturnValue As Boolean = True

        ' All the words in My.Settings.PushBulletTitleFilter must be in TitleIn

        If (TitleIn.Length > 0) AndAlso (My.Settings.PushBulletTitleFilter.Length > 0) Then

            TitleIn = TitleIn.Trim.ToUpper

            Dim TitleInSettings() As String = My.Settings.PushBulletTitleFilter.Trim.ToUpper.Split(" ")

            For Each word In TitleInSettings
                If TitleIn.Contains(word) Then
                Else
                    ReturnValue = False
                    Exit For
                End If
            Next

        Else

            ReturnValue = False

        End If


        'v 4.7 

        ' if  My.Settings.PushBulletTitleFilter contains a "," then look for an exact match of the TitleIn within the filter
        ' this allows for, as an example, "Push2Run W11-Server, From Alexa"

        If (TitleIn.Length > 0) AndAlso (My.Settings.PushBulletTitleFilter.Length > 0) Then

            If (My.Settings.PushBulletTitleFilter.Contains(",")) Then

                Dim TitleU As String = TitleIn.ToUpper.Trim

                Dim FilterU As String = My.Settings.PushBulletTitleFilter.ToUpper.Trim

                If FilterU.Contains(TitleU) Then
                    ReturnValue = True
                End If

            End If

        End If

        Return ReturnValue

    End Function

    Dim objlock1 As Object = New Object
    Private Sub ActionIncomingMessage(ByVal Source As MessageSource, ByVal Message As String, Optional ByVal Description As String = "")

        SyncLock (objlock1)

            Dim ToastTime As DateTime = Now

            Try

                If (Source = MessageSource.Pushbullet) OrElse (Source = MessageSource.Pushover) OrElse (Source = MessageSource.Dropbox) Then

                    If CompetingPush(Source, Message) Then

                        Log("Ignoring competing command")
                        Exit Try

                    End If

                End If

                Dim Request As String = String.Empty
                Dim RunResults As String = String.Empty
                Dim SourceText As String = String.Empty

                gSendingKeyesIsRequired = False

                Select Case Source

                    Case Is = MessageSource.CommandLine
                        SourceText = "Command line"

                    Case Is = MessageSource.Dropbox
                        SourceText = "Dropbox"

                    Case Is = MessageSource.MQTT
                        SourceText = "MQTT"

                    Case Is = MessageSource.Pushbullet
                        SourceText = "Pushbullet"

                    Case Is = MessageSource.Pushover
                        SourceText = "Pushover"

                    Case Is = MessageSource.UserRequest
                        SourceText = "User"

                    Case Else

                        Exit Sub

                End Select

                If Source = MessageSource.UserRequest Then
                    Request = Description
                Else
                    Request = Message
                End If

                Dim ProcessingStatusSoFar As String = SourceText & " requested to run  - """ & Request & """"

                UpdateTheStatusBar(ProcessingStatusSoFar)

                Dim WhatHappened As ActionStatus = ActionStatus.Unknown

                If Description.Length > 0 Then

                    Dim ProgramToRun As String = gCurrentlySelectedRow.Open

                    Dim WorkingDirectory As String = String.Empty
                    If gCurrentlySelectedRow.StartIn.Length > 0 Then WorkingDirectory = gCurrentlySelectedRow.StartIn Else WorkingDirectory = String.Empty

                    Dim Parameters As String = String.Empty

                    If gCurrentlySelectedRow.Parameters.Length > 0 Then Parameters = gCurrentlySelectedRow.Parameters Else Parameters = String.Empty

                    Dim Admin As Boolean = gCurrentlySelectedRow.Admin

                    Dim WindowProcessingStyle As ProcessWindowStyle = gCurrentlySelectedRow.StartingWindowState.ConvertStartingWindowStateToAProcessWindowStyle

                    Dim KeysToSend As String

                    If gCurrentlySelectedRow.KeysToSend.Length > 0 Then
                        KeysToSend = gCurrentlySelectedRow.KeysToSend
                    Else
                        KeysToSend = String.Empty
                    End If

                    WhatHappened = RunProgram(ProgramToRun, WorkingDirectory, Parameters, Admin, WindowProcessingStyle, KeysToSend)

                Else

                    If gMasterStatus = MonitorStatus.Running Then
                        WhatHappened = ActionIncomingMessageNow(Message) '  <************************** go to heart of the program
                    Else
                        WhatHappened = ActionStatus.MasterSwitchWasOff
                    End If

                End If


                Select Case WhatHappened

                    Case Is = ActionStatus.Succeeded
                        RunResults = "Action completed successfully"

                    Case Is = ActionStatus.PartiallySucceeded
                        RunResults = "Multiple actions completed partially successfully"

                    Case Is = ActionStatus.Failed
                        RunResults = "The action was run, but it appears to have failed"

                    Case Is = ActionStatus.NotProcessedAsNoMatchingPhrasesFound
                        RunResults = "No matching phrases found"

                    Case Is = ActionStatus.NotProcessedWhileAtLeastOneMatchingPhraseWasFoundNoneWereEnabled
                        RunResults = "No matching enabled phrases found"

                    Case Is = ActionStatus.NotProcecessAsAUACPromptWouldBeRequired
                        RunResults = "Program not run as a UAC prompt would be required"

                    Case Is = ActionStatus.MasterSwitchWasOff
                        RunResults = "Command was not actioned as the Master Switch is off"

                    Case Is = ActionStatus.Unknown
                        RunResults = "Results unverified"

                    Case Is = ActionStatus.LeaveBlank
                        RunResults = ""

                End Select

                ProcessingStatusSoFar &= " - " & RunResults
                UpdateTheStatusBar(ProcessingStatusSoFar, WhatHappened)

                Log(RunResults)
                Log("")

                If Me.Visibility = Visibility.Visible Then  ' v4.9
                    Thread.Sleep(500)
                End If

                If My.Settings.ShowNotifications Then

                    Dim Timeout As Integer = 0
                    While (gSendingKeyesIsRequired) AndAlso (Timeout < 1000)
                        Thread.Sleep(10)
                        Timeout += 1
                        DoEvents()
                    End While

                    ToastNotification(Request, Capitalize(RunResults), SourceText & " request", ToastTime)

                End If

            Catch ex As Exception

            End Try

        End SyncLock

    End Sub

    Private Function UpdateFieldWithRegexSubstitutions(ByVal SpeechInput As String, ByVal ListenFor As String, ByVal FieldToUpdate As String) As String

        Dim ReturnValue As String = FieldToUpdate

        Try

            Dim RegexRule = New Regex(ListenFor, RegexOptions.IgnoreCase)

            Dim Match As Match = RegexRule.Match(SpeechInput)

            If Match.Success Then

                For Each groupName As String In RegexRule.GetGroupNames()

                    ReturnValue = ReturnValue.Replace("<" + groupName + ">", Match.Groups(groupName).Value)

                Next

            End If

        Catch ex As Exception

            ReturnValue = FieldToUpdate

        End Try

        Return ReturnValue

    End Function

    Private Function Capitalize(ByVal input As String)

        Dim ReturnValue As String = String.Empty

        If input.Length > 0 Then
            ReturnValue = input.Remove(1).ToUpper & input.Remove(0, 1)
        End If

        Return ReturnValue

    End Function

    Private Structure HistoryEntry
        Dim Source As MessageSource
        Dim Message As String
        Dim TimeStamp As DateTime
    End Structure

    Dim History As New List(Of HistoryEntry)

    Private Function CompetingPush(ByVal Source As MessageSource, ByVal Message As String) As Boolean

        Dim MatchFound As Boolean = False

        ClearHistoryBasedOnTime()

        'look for a matching history entry from another source

        For Each entry In History

            If (Message = entry.Message) AndAlso (Source <> entry.Source) Then
                MatchFound = True
                Exit For
            End If

        Next

        If MatchFound Then
        Else
            'add newest entry to list
            Dim NewHistoryEntry As New HistoryEntry
            With NewHistoryEntry
                .Source = Source
                .Message = Message
                .TimeStamp = Now
            End With

            History.Add(NewHistoryEntry)
        End If

        Return MatchFound

    End Function

    Private Sub ClearHistoryBasedOnTime()

        ' My.Settings.CompetingPushThreshold is not yet set in the options window, rather its value is defaulted to -60 (seconds) in the settings file
        ' it can be changed there until there is actually a way to have ifttt send pushes from pushbullet and pushover at the same time

        SyncLock (History)

            For Each entry In History.ToList() ' .ToList is used because an item from the history list can't be removed while it is being iterated

                If entry.TimeStamp < Now.AddSeconds(-1 * My.Settings.CompetingPushThreshold) Then
                    History.Remove(entry)
                End If

            Next

        End SyncLock

    End Sub


    Private Sub Websocket_DataReceived_Pushbullet(ByVal sender As Object, ByVal e As WebSocket4Net.DataReceivedEventArgs)
        'Log("Data Received")
        'Log("")
    End Sub

    Private WebSocketErrorWasReportedPushbullet As Boolean = False

    Private Sub Websocket_Error_Pushbullet(ByVal sender As Object, ByVal e As SuperSocket.ClientEngine.ErrorEventArgs)

        WebSocketErrorWasReportedPushbullet = True

        Log("Pushbullet websocket error")
        Log(e.Exception.Message)

    End Sub

    Private Sub Websocket_Closed_Pushbullet(ByVal sender As Object, ByVal e As EventArgs)

        Log("Pushbullet websocket closed")
        Log("")

        If WebSocketErrorWasReportedPushbullet Then
            ResetTheTimePushbulletWasLastAccessed(Now.AddDays(-1))
        End If

    End Sub

    Private Sub Websocket_DataReceived_Pushover(ByVal sender As Object, ByVal e As WebSocket4Net.DataReceivedEventArgs)

        If DisablePushoverProcessing Then Exit Sub

        Dim PushoverData As String = System.Text.Encoding.UTF8.GetString(e.Data)

        If gCriticalPushOverErrorReported Then
        Else

            Select Case PushoverData

                Case Is = "#"
                    LastTimeDataWasReceivedFromPushover = Now
                ' just saying the connection is live

                Case Is = "!"
                    LastTimeDataWasReceivedFromPushover = Now
                    ProcessANewPushoverMessage()

                Case Is = "R"
                    LastTimeDataWasReceivedFromPushover = Now
                    Log("Pushover has requested a reconnect")
                    Log("Reconnect underway ...")
                    OpenThePushoverWebSocket()
                    Log("Reconnect complete")
                    Log("")

                Case Is = "E"

                    gCriticalPushOverErrorReported = True

                    LastTimeDataWasReceivedFromPushover = Now.AddDays(-1)
                    Log("Pushover has reported that something is wrong - but not exactly what")
                    Log("The problem is likely with your Pushover ID and/or password or your Pushover device")
                    Log("You may want to try (re)authenticating Pushover in the Actions - Options - Pushover window")
                    Log("")
                    Log("Pushover monitoring will be discontinued")
                    Log("")

                    CloseThePushoverWebSocket()

            End Select

        End If

    End Sub

    Private Sub Websocket_Error_Pushover(ByVal sender As Object, ByVal e As SuperSocket.ClientEngine.ErrorEventArgs)

        Log("Pushover websocket error")
        Log(e.Exception.Message)

    End Sub

    Private Sub Websocket_Closed_Pushover(ByVal sender As Object, ByVal e As EventArgs)

        Log("Pushover websocket closed")
        Log("")

        ResetTheTimePushoverWasLastAccessed(Now.AddDays(-1))

    End Sub

    Private Function GetServerTimeOfMostRecentPush_Unix() As Double

        Dim ReturnValue As Double = 0

        Try

            Dim ServerResponse As String = String.Empty

            SendRequest(EncryptionClass.Decrypt(My.Settings.PushBulletAPI), "GET", AddressForGettingPushes & "?limit=1", String.Empty, ServerResponse)

            ReturnValue = CType(GetFirstMatchingValueFromJSONResponseString("modified", ServerResponse), Double)

        Catch ex As Exception

            ReturnValue = GetUnixTime(Now.AddDays(-1)) ' could not establish time of last post, so use 1 days ago as the basis for the response

        End Try

        Return ReturnValue

    End Function

    'Private Function GetValueFromMostRecentPost(ByVal KeyToLookFor As String, ByVal ServerResponse As String) As String

    '    'Dim settingValuesFromPreviousSettingsFile As String = String.Empty

    '    'Try

    '    '    Dim Responses() As String = ServerResponse.Replace("""", "").Split(",")

    '    '    For Each IndividualResponse In Responses
    '    '        If IndividualResponse.StartsWith(KeyToLookFor) Then
    '    '            settingValuesFromPreviousSettingsFile = IndividualResponse.Remove(0, KeyToLookFor.Length + 1)
    '    '            settingValuesFromPreviousSettingsFile = settingValuesFromPreviousSettingsFile.TrimEnd("]").TrimEnd("}")
    '    '            Exit For
    '    '        End If
    '    '    Next

    '    'Catch ex As Exception

    '    'End Try

    '    'Return settingValuesFromPreviousSettingsFile

    'End Function

    'testing moved to common 

    'Private Shared Function GetFirstMatchingValueFromJSONResponseString(ByVal KeyToLookFor As String, ByVal ServerResponse As String) As String


    '    On Error Resume Next ' changed try catch to on error resume next in 2.4 - to work around problem returning responses involving boolean values

    '    If ServerResponse.Length > 0 Then

    '        Dim DictionaryOfJSONResults = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(ServerResponse)

    '        For Each item In DictionaryOfJSONResults

    '            Dim ItemKey As String = item.Key
    '            Dim ItemValue As JArray = item.Value

    '            If ItemKey = KeyToLookFor Then

    '                If ItemValue Is Nothing Then
    '                    Return CType(item.Value, String)
    '                Else
    '                    Return item.Value.ToString ' changed in v2.5 from ItemValue.ToString
    '                End If

    '            ElseIf ItemValue.HasValues Then

    '                For Each child In ItemValue

    '                    For Each ChildProperty As JProperty In child.Children

    '                        If ChildProperty.Name = KeyToLookFor Then

    '                            Return ChildProperty.Value.ToString

    '                        End If

    '                    Next

    '                Next

    '            End If

    '        Next

    '    End If

    '    Return String.Empty

    'End Function



    Private Function GetNotesSinceLastRequest() As NoteInfo

        Dim ReturnValue As New NoteInfo
        With ReturnValue
            .Title = String.Empty
            .Body = String.Empty
        End With

        Try

            Dim ServerResponse As String = String.Empty

            'v1.2 added .Replace(",", ".") to deal with countries that use a comma instead of a period in their date format

            SendRequest(EncryptionClass.Decrypt(My.Settings.PushBulletAPI), "GET", AddressForGettingPushes & "?modified_after=" & LastTimeAPushbulletPushWasRecieved_Unix.ToString.Replace(",", "."), String.Empty, ServerResponse)

            ReturnValue.Iden = GetFirstMatchingValueFromJSONResponseString("iden", ServerResponse)
            ReturnValue.Title = GetFirstMatchingValueFromJSONResponseString("title", ServerResponse)
            ReturnValue.Body = GetFirstMatchingValueFromJSONResponseString("body", ServerResponse)
            ReturnValue.Dismissed = CType(GetFirstMatchingValueFromJSONResponseString("dismissed", ServerResponse), Boolean)

            LastTimeAPushbulletPushWasRecieved_Unix = CType(GetFirstMatchingValueFromJSONResponseString("modified", ServerResponse), Double)

        Catch ex As Exception

        End Try

        Return ReturnValue

    End Function

    Private Sub SendALink(ByVal APIKey As String, ByVal Title As String, ByVal Body As String, ByVal url As String)

        Dim DataString As String = String.Format("{{ ""type"": ""{0}"", ""title"": ""{1}"", ""body"": ""{2}"", ""url"": ""{3}"" }}", "link", Title, Body, url)
        SendRequest(APIKey, "POST", AddressForSendingPushes, DataString)

    End Sub

    Private Sub SendANote(ByVal APIKey As String, ByVal Title As String, ByVal Body As String)

        Dim DataString As String = String.Format("{{ ""type"": ""{0}"", ""title"": ""{1}"", ""body"": ""{2}"" }}", "note", Title, Body)
        SendRequest(APIKey, "POST", AddressForSendingPushes, DataString)

    End Sub

    Private Sub DissmissANote(ByVal APIKey As String, ByVal iden As String)

        Dim ServerAddress As String = AddressForDismissingPushes.Replace("{iden}", iden)
        SendRequest(APIKey, "POST", ServerAddress, "{ ""dismissed"": true }")

    End Sub

    Private Function KeepPushbulletAccountActive(ByVal APIKey As String) As Boolean

        Dim ReturnValue As Boolean = False

        Try

            Dim UserIden As String = String.Empty
            Dim ServerResponse As String = String.Empty

            SendRequest(APIKey, "GET", AddressForGettingPushbulletUserIden, String.Empty, ServerResponse)
            UserIden = GetFirstMatchingValueFromJSONResponseString("iden", ServerResponse)

            Dim ServerAddress As String = AddressForKeepingPushbulletAccountActive
            Dim DataToPost As String = "{ ""name"": ""push2run_active"", ""user_iden"": ""client_user_iden"" }".Replace("client_user_iden", UserIden)

            ReturnValue = SendRequest("", "POST", ServerAddress, DataToPost)

        Catch ex As Exception

        End Try

        Return ReturnValue

    End Function

    'Private Function GetDevices(ByVal APIKey As String) As String

    '    Dim Response As String = String.Empty
    '    SendRequest(APIKey, "GET", AddressForGettingDevices, String.Empty, Response)
    '    Return Response

    'End Function

    Private Function GetUnixTime(ByVal TimeToConvert As DateTime) As Double
        Return (TimeToConvert.UtcNow - New DateTime(1970, 1, 1, 0, 0, 0)).TotalSeconds
    End Function

#End Region

#Region "Filewatcher"

    Private Const gFileWatcherFilterName As String = "command*.txt"
    Private gFileWatcherPathName As String = String.Empty
    Private gFileWatcherPathAndFileName As String = String.Empty

    Friend UniqueSessionID As String = String.Empty

    Friend Const CommandToOpenUpMainWindow As String = "open the Push2Run main window"

    Private Sub InitializeFileWatcherVariables()

        Randomize()
        UniqueSessionID = (Rnd() * 10000000).ToString

        ' Enable Push2Run to watch for commands from other instances of Push2Run

        gFileWatcherPathName = System.IO.Path.GetTempPath() & gThisProgramName
        gFileWatcherPathAndFileName = gFileWatcherPathName & "\" & gFileWatcherFilterName

        If Directory.Exists(gFileWatcherPathName) Then
        Else
            Directory.CreateDirectory(gFileWatcherPathName)
        End If

        If File.Exists(gFileWatcherPathAndFileName) Then
            System.IO.File.Delete(gFileWatcherPathAndFileName)
        End If

    End Sub

    Private Sub SetupForFileWatcherMonitoring()

        Try

            'Delete any files that may be in the directory at startup
            For Each deleteFile In Directory.GetFiles(gFileWatcherPathName, gFileWatcherFilterName, SearchOption.TopDirectoryOnly)
                File.Delete(deleteFile)
            Next

            Dim fw As New FileSystemWatcher
            fw.Path = gFileWatcherPathName
            fw.IncludeSubdirectories = False
            fw.Filter = gFileWatcherFilterName
            AddHandler fw.Created, New FileSystemEventHandler(AddressOf FileWatcherDetectedAFileWasCreated)
            AddHandler fw.Error, New ErrorEventHandler(AddressOf FileWatcherError)
            fw.EnableRaisingEvents = True

        Catch ex As Exception
            Log("Problem setting up command line monitoring: " & vbCrLf & gFileWatcherPathName & vbCrLf & ex.Message.ToString)  'v4.2  (try to figure why this is thrown on install)
        End Try

    End Sub


    Private LockedObject As New Object
    Friend Sub FileWatcherDetectedAFileWasCreated(ByVal source As Object, ByVal e As FileSystemEventArgs)

        SyncLock LockedObject

            Try

                Dim DetectedFileName As String = gFileWatcherPathName & "\" & e.Name

                Dim IncomingData As String = String.Empty

                Dim MaxNumberOfSecondsToSpendWaitingForFileToBecomeAvailable As Decimal = 5.0
                Dim NumberOfSecondsSpentWaitingForFileToBecomeAvailable As Decimal = 0

                While (NumberOfSecondsSpentWaitingForFileToBecomeAvailable < MaxNumberOfSecondsToSpendWaitingForFileToBecomeAvailable)

                    Try

                        IncomingData = File.ReadAllText(DetectedFileName)

                        NumberOfSecondsSpentWaitingForFileToBecomeAvailable = MaxNumberOfSecondsToSpendWaitingForFileToBecomeAvailable

                    Catch ex As Exception
                        System.Threading.Thread.Sleep(100)
                        NumberOfSecondsSpentWaitingForFileToBecomeAvailable += 0.1
                    End Try

                End While

                If IncomingData.Contains(vbCr) Then

                    Dim DataReceived() = IncomingData.Split(vbCr)

                    If DataReceived.Count = 2 Then

                        If DataReceived(0) = UniqueSessionID Then

                            ' file was created by this instance, ignor it

                        Else

                            If DataReceived(1) = CommandToOpenUpMainWindow Then

                                Me.Dispatcher.Invoke(New OpenMainWindowCallback(AddressOf OpenMainWindow))

                            Else

                                Dim DataIn As String = DataReceived(1)

                                If DataIn.ToUpper.EndsWith(".P2R""") Then

                                    'v4.2
                                    Try

                                        Log("Importing: " & DataIn)
                                        Dim ImportFileName As String = DataIn.Trim("""")
                                        Dim FileContents As String = File.ReadAllText(ImportFileName)

                                        If FileContents.StartsWith("<?xml") Then
                                            Dim LoadedCard As CardClass = LoadACard(ImportFileName)

                                            With gCurrentlySelectedRow

                                                .SortOrder = 15

                                                .Description = LoadedCard.Description
                                                .ListenFor = LoadedCard.ListenFor
                                                .Open = LoadedCard.Open
                                                .StartIn = LoadedCard.StartDirectory
                                                .Parameters = LoadedCard.Parameters
                                                .Admin = LoadedCard.StartWithAdminPrivileges
                                                .StartingWindowState = LoadedCard.StartingWindowState
                                                .KeysToSend = LoadedCard.KeysToSend

                                            End With

                                            SafelyAddgCurrentlySelectedRecordIntoTheDatabase(False, ImportFileName)

                                            GetAddChangeInfo("Change")

                                        Else
                                            gImportFileName = ImportFileName
                                            DoImport()
                                        End If

                                    Catch ex As Exception

                                    End Try


                                Else
                                    Log("Incoming command line request ...")
                                    Log(DataIn)
                                    ActionIncomingMessage(MessageSource.CommandLine, DataIn)
                                End If


                            End If

                        End If

                    End If

                End If

                System.IO.File.Delete(gFileWatcherPathName & "\" & e.Name)

                System.Threading.Thread.Sleep(1000) ' give time for the file to be deleted

            Catch ex As Exception

                Log("Problem with command line monitoring: " & ex.ToString)

            End Try

        End SyncLock

    End Sub

    Friend Sub FileWatcherError(ByVal source As Object, ByVal e As System.IO.ErrorEventArgs)

        Try

            Log("File watcher error: " & e.ToString)

        Catch ex As Exception

        End Try

    End Sub

    Friend Sub CreateACommandToBePickedUpByFileWatcher(ByVal Command As String)

        If Command = String.Empty Then Exit Sub

        Try

            Command = UniqueSessionID & vbCr & Command

            Dim FileToCreate As String = gFileWatcherPathAndFileName.Replace("*", Format(Now, "-yyyy-MM-dd-hh-mm-ss-fff"))

            System.IO.File.WriteAllText(FileToCreate, Command)
            System.Threading.Thread.Sleep(250)

        Catch ex As Exception

            Log("Problem creating a command to send to another instance of Push2Run: " & ex.ToString)

        End Try

    End Sub

#End Region
    Private Sub RunCurrentlySelectedRecord()

        Dim DescriptionToUse As String = String.Empty

        Dim LogMessage As String = String.Empty

        If gCurrentlySelectedRow.Description.Length > 0 Then

            DescriptionToUse = gCurrentlySelectedRow.Description
            LogMessage = "Client requested to run the currently selected entry, description is " & DescriptionToUse

        Else

            DescriptionToUse = "the currently selected entry, program = " & gCurrentlySelectedRow.Open
            LogMessage = "Client requested to run the currently selected entry, program is " & gCurrentlySelectedRow.Open

        End If

        'v1.6 prompt when there is a $ sign in the open or parm fields

        Dim SimulateSpeechInput As String = String.Empty

        Dim HoldCurrentlySelectedRow As gRowRecord = gCurrentlySelectedRow

        ' as the $ sign is use as the variable text this is a little dodge to keep the $ sign in the KeysToSend field
        ' if the $ sign is needed, then the user needs to use {$} otherwise the $ sign by itself will be used as the variable text marker

        Dim HighValueString As String = Chr(255).ToString

        If gCurrentlySelectedRow.KeysToSend.Contains("{$}") Then
            gCurrentlySelectedRow.KeysToSend = Microsoft.VisualBasic.Strings.Replace(gCurrentlySelectedRow.KeysToSend, "{$}", HighValueString, , , CompareMethod.Text)
        End If

        Dim ReplacementIsRequiredForOpen As Boolean = gCurrentlySelectedRow.Open.Contains("$")
        Dim ReplacementIsRequiredForParms As Boolean = gCurrentlySelectedRow.Parameters.Contains("$")
        Dim ReplacementIsRequiredForKeysToSend As Boolean = gCurrentlySelectedRow.KeysToSend.Contains("$")

        Dim ReplacementRequiredAsListenToFieldIsUsingRegex As Boolean = False

        Try
            Dim RegexRule = New Regex(gCurrentlySelectedRow.ListenFor, RegexOptions.IgnoreCase)
            ReplacementRequiredAsListenToFieldIsUsingRegex = (RegexRule.GetGroupNames.Count > 1)
        Catch ex As Exception
        End Try

        If ReplacementIsRequiredForOpen OrElse ReplacementIsRequiredForParms OrElse ReplacementIsRequiredForKeysToSend Then

            MakeTopMost(SafeNativeMethods.FindWindow(Nothing, Me.Title), False) 'v2.0.1 applied correction here
            SimulateSpeechInput = InputBox("Please enter what you would like to use as the variable text (the part that replaces the $ sign)", gThisProgramName, "", MouseDownMousePosition.X, MouseDownMousePosition.Y)
            MakeTopMost(SafeNativeMethods.FindWindow(Nothing, Me.Title), My.Settings.AlwaysOnTop) 'v2.0.1 applied correction here

        ElseIf ReplacementRequiredAsListenToFieldIsUsingRegex Then

            MakeTopMost(SafeNativeMethods.FindWindow(Nothing, Me.Title), False) 'v2.0.1 applied correction here
            SimulateSpeechInput = InputBox("Please enter everything you would say to run this card", gThisProgramName, "", MouseDownMousePosition.X, MouseDownMousePosition.Y)
            MakeTopMost(SafeNativeMethods.FindWindow(Nothing, Me.Title), My.Settings.AlwaysOnTop) 'v2.0.1 applied correction here

        End If

        If SimulateSpeechInput > String.Empty Then

            'v4.4

            'If ReplacementIsRequiredForOpen Then gCurrentlySelectedRow.Open = gCurrentlySelectedRow.Open.Replace("$", SimulateSpeechInput)

            'If ReplacementIsRequiredForParms Then gCurrentlySelectedRow.Parameters = gCurrentlySelectedRow.Parameters.Replace("$", SimulateSpeechInput)

            If ReplacementIsRequiredForOpen Then
                SpecialSubstitutions(gCurrentlySelectedRow.Open, "", SimulateSpeechInput, gCurrentlySelectedRow.Open, "")
                gCurrentlySelectedRow.Open = gCurrentlySelectedRow.Open.Replace("$", SimulateSpeechInput)
            End If

            If ReplacementIsRequiredForParms Then
                SpecialSubstitutions("", gCurrentlySelectedRow.Parameters, SimulateSpeechInput, "", gCurrentlySelectedRow.Parameters)
                gCurrentlySelectedRow.Parameters.Replace("$", SimulateSpeechInput)
            End If

            '**

            If ReplacementIsRequiredForKeysToSend Then gCurrentlySelectedRow.KeysToSend = Microsoft.VisualBasic.Strings.Replace(gCurrentlySelectedRow.KeysToSend, "$", SimulateSpeechInput, , , CompareMethod.Text)

            gCurrentlySelectedRow.Open = UpdateFieldWithRegexSubstitutions(SimulateSpeechInput, gCurrentlySelectedRow.ListenFor, gCurrentlySelectedRow.Open)

            gCurrentlySelectedRow.Parameters = UpdateFieldWithRegexSubstitutions(SimulateSpeechInput, gCurrentlySelectedRow.ListenFor, gCurrentlySelectedRow.Parameters)

            gCurrentlySelectedRow.KeysToSend = UpdateFieldWithRegexSubstitutions(SimulateSpeechInput, gCurrentlySelectedRow.ListenFor, gCurrentlySelectedRow.KeysToSend)


        End If

        If gCurrentlySelectedRow.KeysToSend.Contains(HighValueString) Then
            gCurrentlySelectedRow.KeysToSend = Microsoft.VisualBasic.Strings.Replace(gCurrentlySelectedRow.KeysToSend, HighValueString, "$", , , CompareMethod.Text)
        End If

        If (SimulateSpeechInput.Length = 0) AndAlso ((ReplacementIsRequiredForOpen OrElse ReplacementIsRequiredForParms OrElse ReplacementIsRequiredForKeysToSend)) Then ' OrElse RegExReplacementIsRequiredForOpen OrElse RegExReplacementIsRequiredForParms) Then

            Beep()
            Call TopMostMessageBox(gCurrentOwner, "Variable text is required to running this card.", "Push2Run - Warning", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK)

        Else

            Log(LogMessage)
            ActionIncomingMessage(MessageSource.UserRequest, "", DescriptionToUse)

        End If

        gCurrentlySelectedRow = HoldCurrentlySelectedRow

    End Sub

    Private Function ActionIncomingMessageNow(ByVal IncomingMessage As String) As ActionStatus

        'v3.6 added to allow multiple commands to run, each separated by Separating words

        Dim ReturnValue As ActionStatus = ActionStatus.NotYetSet

        Try

            Dim RedoRequired As Boolean = False

            IncomingMessage = IncomingMessage.ToLower.Trim
            Dim RemainingMessage As String = IncomingMessage

            ' if Separating words are not to be used run everything as was done prior to version 3.6

            If gUseSeparatingWords Then

                Dim SeparatingWordsInIncommingMessage As Boolean = False

                ' note: by the time they get here Separating words are all in lower case
                For Each SeparatingWord In gSeparatingWords
                    If IncomingMessage.Contains(SeparatingWord) Then
                        SeparatingWordsInIncommingMessage = True
                        Exit For
                    End If
                Next

                If SeparatingWordsInIncommingMessage Then
                Else
                    ReturnValue = ActionIncomingIndivitualMessageNow(IncomingMessage, RedoRequired, True)
                    Exit Try
                End If

                ' ok, Separating words are needed and present, so lets deal with them

                For Each SeparatingWord In gSeparatingWords
                    IncomingMessage = IncomingMessage.Replace(" " & SeparatingWord & " ", Chr(255))
                Next

                Dim SplitMessages() As String = IncomingMessage.Split(Chr(255))

                For Each IndividualMessage In SplitMessages

                    IndividualMessage = IndividualMessage.Trim

                    If IndividualMessage.Length > 0 Then

                        Dim IndividualReturnValue As ActionStatus = ActionIncomingIndivitualMessageNow(IndividualMessage, RedoRequired, False)

                        If RedoRequired Then
                            IndividualReturnValue = ActionIncomingIndivitualMessageNow(RemainingMessage, RedoRequired, True)
                            RemainingMessage = String.Empty
                        End If

                        With IndividualReturnValue

                            Select Case IndividualReturnValue

                                Case .MasterSwitchWasOff

                                    ReturnValue = ActionStatus.MasterSwitchWasOff

                                Case .Succeeded

                                    If (ReturnValue = ActionStatus.NotYetSet) OrElse (ReturnValue = ActionStatus.Succeeded) Then
                                        ReturnValue = ActionStatus.Succeeded
                                    Else
                                        ReturnValue = ActionStatus.PartiallySucceeded
                                    End If

                                    'the following two lines give time for the windows state to be properly set on the program that is loading
                                    Thread.Sleep(1500)

                                Case Else ' .Failed, .NotProcecessAsAUACPromptWouldBeRequired, .NotProcecessNoProgramToRun, .NotProcessedAsNoMatchingPhrasesFound, .NotProcessedWhileAtLeastOneMatchingPhraseWasFoundNoneWereEnabled

                                    If (ReturnValue = ActionStatus.NotYetSet) Then
                                        ReturnValue = IndividualReturnValue

                                    ElseIf (ReturnValue = ActionStatus.Succeeded) OrElse (ReturnValue = ActionStatus.PartiallySucceeded) Then
                                        ReturnValue = ActionStatus.PartiallySucceeded

                                    ElseIf (ReturnValue = IndividualReturnValue) Then
                                        ' just keep the return value as it is

                                    Else

                                        ReturnValue = ActionStatus.Unknown

                                    End If

                            End Select

                        End With

                        '
                        ' update remaining message to remove the portion of what has already been processed

                        If RemainingMessage.Length > 0 Then

                            Try

                                RemainingMessage = RemainingMessage.Remove(0, IndividualMessage.Length)

                                Dim KeepReducingRemaningMessage As Boolean = True

                                While KeepReducingRemaningMessage

                                    RemainingMessage = RemainingMessage.Trim

                                    KeepReducingRemaningMessage = False

                                    For Each SeparatingWord In gSeparatingWords
                                        If RemainingMessage.StartsWith(SeparatingWord) Then
                                            RemainingMessage = RemainingMessage.Remove(0, SeparatingWord.Length).Trim
                                            KeepReducingRemaningMessage = True
                                            Exit For
                                        End If
                                    Next

                                End While

                            Catch ex As Exception
                            End Try

                        End If

                        If RemainingMessage.Length = 0 Then
                            Exit Try
                        End If

                    End If

                Next

            Else

                ' if there are no Separating words in the incoming message run everything as was done prior to version 3.6

                ReturnValue = ActionIncomingIndivitualMessageNow(IncomingMessage, RedoRequired, True)

            End If

        Catch ex As Exception

        End Try

        Return ReturnValue

    End Function


    Private LastReload As DateTime = Now.AddHours(-1)

    Private Function ActionIncomingIndivitualMessageNow(ByVal IncomingMessage As String, ByRef RedoRequired As Boolean, ByVal MakeItHappen As Boolean) As ActionStatus

        Dim ReturnValue As ActionStatus = ActionStatus.Unknown

        RedoRequired = False

        Try

            Dim SpecialProcessing_NoMatchingPhrases As Boolean = False
            Dim SpecialProcessing_NoMatchingEnabledPhrases As Boolean = False

            If IncomingMessage.Trim.ToLower.StartsWith("no matching phrases") Then
                SpecialProcessing_NoMatchingPhrases = True
            End If

            If IncomingMessage.Trim.ToLower.StartsWith("no matching enabled phrases") Then
                SpecialProcessing_NoMatchingEnabledPhrases = True
            End If

            Dim UnmatchedIncomingMessage As String = IncomingMessage

            If SpecialProcessing_NoMatchingPhrases Then
                If IncomingMessage.Length > "no matching phrases".Length Then
                    UnmatchedIncomingMessage = IncomingMessage.Remove(0, "no matching phrases".Length).Trim
                Else
                    UnmatchedIncomingMessage = String.Empty
                End If
            End If

            If SpecialProcessing_NoMatchingEnabledPhrases Then
                If IncomingMessage.Length > "no matching enabled phrases".Length Then
                    UnmatchedIncomingMessage = IncomingMessage.Remove(0, "no matching enabled phrases".Length).Trim
                Else
                    UnmatchedIncomingMessage = String.Empty
                End If
            End If

            Dim AnActionWasAttempted As Boolean = False

            ' LoadFromDatabase() 'always reload the database v2.5.4  
            ' v4.9  only reload if the database has not been reloaded based on a rolling 20 seconds interval
            If LastReload > Now Then
            Else
                LoadFromDatabase()
            End If
            LastReload = Now.AddSeconds(20)

            Dim FoundMatchingEntryButItWasSwitchedOff As Boolean = False

            Dim ProgramToRun As String = String.Empty
            Dim WorkingDirectory As String = String.Empty
            Dim Parameters As String = String.Empty
            Dim Admin As Boolean = False
            Dim WindowProcessingStyle As ProcessWindowStyle = ProcessWindowStyle.Normal
            Dim KeysToSend As String = String.Empty

            Dim GenericMessage As String = String.Empty

            Dim IgnorSeparatingWords As Boolean = False

            For Each CommandItem As gCommandRecord In CommandTable

                Dim AllPhrasesToListenFor() As String = CommandItem.ListenFor.Split(vbCr)

                For Each SpecificPhraseToListenFor As String In AllPhrasesToListenFor

                    SpecificPhraseToListenFor = SpecificPhraseToListenFor.Trim

                    IgnorSeparatingWords = SpecificPhraseToListenFor.Contains("$")

                    SpecificPhraseToListenFor = SpecificPhraseToListenFor.Trim

                    'v4.2

                    Dim RegexMatch As Boolean

                    'vx.x  recoded to exclude a regex match if the incoming message just starts with a specific 'Listen for' phrase; allow for mqtt topics

                    Try

                        If (My.Settings.UseMQTT AndAlso (Not My.Settings.MQTTListenForPayloadOnly) AndAlso IncomingMessage.Contains("/")) Then

                            ' assume there is an mqtt topic and strip it (the topic) out of the incoming message for testing purposes

                            Dim TopicRemovedIncomingMessage As String = IncomingMessage.Remove(0, IncomingMessage.LastIndexOf("/") + 1).Trim
                            If TopicRemovedIncomingMessage.ToUpper = SpecificPhraseToListenFor.ToUpper Then
                                Exit Try
                            End If

                        Else

                            If IncomingMessage.ToUpper = SpecificPhraseToListenFor.ToUpper Then
                                Exit Try
                            End If

                        End If

                        RegexMatch = (Regex.Matches(IncomingMessage.Trim, SpecificPhraseToListenFor.Trim, RegexOptions.IgnoreCase).Count > 0)

                    Catch ex As Exception

                        RegexMatch = False

                    End Try

                    If (Not SpecialProcessing_NoMatchingPhrases AndAlso Not SpecialProcessing_NoMatchingEnabledPhrases) AndAlso
                       ((IncomingMessage.Trim.ToUpper = SpecificPhraseToListenFor.Trim.ToUpper) OrElse
                       (IncomingMessage.Trim.ToUpper Like SpecificPhraseToListenFor.Trim.ToUpper) OrElse ' v3.4.2 added to allow the like command
                       (IncomingMessage.Trim.ToUpper Like SpecificPhraseToListenFor.Trim.ToUpper.Replace("* ", "*").Replace(" *", "*")) OrElse   ' v3.6 added to improve the use of the like command
                       RegexMatch) Then 'v4.1

                        If CommandItem.DesiredStatus = StatusValues.SwitchOn Then

                            ProgramToRun = CommandItem.Open

                            If CommandItem.StartIn.Length > 0 Then WorkingDirectory = CommandItem.StartIn Else WorkingDirectory = String.Empty

                            If CommandItem.Parameters.Length > 0 Then Parameters = CommandItem.Parameters Else Parameters = String.Empty

                            Admin = CommandItem.Admin

                            WindowProcessingStyle = CommandItem.StartingWindowState.ConvertStartingWindowStateToAProcessWindowStyle

                            KeysToSend = CommandItem.KeysToSend

                            AnActionWasAttempted = True

                            '*************************************************************************************************************
                            'v4.2

                            ProgramToRun = UpdateFieldWithRegexSubstitutions(IncomingMessage, SpecificPhraseToListenFor, ProgramToRun)
                            Parameters = UpdateFieldWithRegexSubstitutions(IncomingMessage, SpecificPhraseToListenFor, Parameters)
                            KeysToSend = UpdateFieldWithRegexSubstitutions(IncomingMessage, SpecificPhraseToListenFor, KeysToSend)

                            '*************************************************************************************************************

                            If MakeItHappen Then
                                ReturnValue = RunProgram(ProgramToRun, WorkingDirectory, Parameters, Admin, WindowProcessingStyle, KeysToSend)
                            Else
                                If IgnorSeparatingWords Then
                                    RedoRequired = True
                                    ReturnValue = ActionStatus.Unknown
                                    GoTo EarlyOut
                                Else
                                    ReturnValue = RunProgram(ProgramToRun, WorkingDirectory, Parameters, Admin, WindowProcessingStyle, KeysToSend)
                                End If
                            End If

                            Exit For

                            '*************************************************************************************************************

                        Else

                            FoundMatchingEntryButItWasSwitchedOff = True

                            If IgnorSeparatingWords Then
                                RedoRequired = True
                            End If

                        End If

                    End If

                    ' test for the case where the listen for phrase ends with a $

                    If SpecificPhraseToListenFor.EndsWith("$") OrElse SpecialProcessing_NoMatchingPhrases OrElse SpecialProcessing_NoMatchingEnabledPhrases Then

                        If SpecificPhraseToListenFor.EndsWith("$") Then
                            GenericMessage = Microsoft.VisualBasic.Left(SpecificPhraseToListenFor, SpecificPhraseToListenFor.Length - 1).Trim ' remove the $ (variable part of the phrase)
                        Else
                            GenericMessage = Microsoft.VisualBasic.Left(SpecificPhraseToListenFor, SpecificPhraseToListenFor.Trim.Length).Trim
                        End If

                        If GenericMessage.Length > 0 Then

                            If (IncomingMessage.ToUpper.Trim.StartsWith(GenericMessage.ToUpper)) OrElse (IncomingMessage.ToUpper.Trim Like GenericMessage.ToUpper) OrElse ' Then   ' v3.4.2 testing here to allow the like command
                               (IncomingMessage.Trim.ToUpper Like SpecificPhraseToListenFor.Trim.ToUpper.Replace("$ ", "*").Replace(" $", "*").Replace("$", "*")) Then  ' v3.6.1 testing here

                                Dim VariablePortionOfIncomingMessage As String = "" 'v3.6.1

                                If GenericMessage.Contains("*") Then

                                    Try
                                        Dim modifiedGenericMessage = GenericMessage.Replace("*", "").Trim
                                        VariablePortionOfIncomingMessage = IncomingMessage.Remove(0, IncomingMessage.LastIndexOf(modifiedGenericMessage) + modifiedGenericMessage.Length).Trim
                                    Catch ex As Exception
                                    End Try

                                Else
                                    VariablePortionOfIncomingMessage = IncomingMessage.Remove(0, GenericMessage.Length).Trim
                                End If

                                If SpecialProcessing_NoMatchingPhrases OrElse SpecialProcessing_NoMatchingEnabledPhrases Then
                                    VariablePortionOfIncomingMessage = UnmatchedIncomingMessage
                                End If

                                SpecialSubstitutions(CommandItem.Open, CommandItem.Parameters, VariablePortionOfIncomingMessage, ProgramToRun, Parameters)

                                If CommandItem.StartIn.Length > 0 Then WorkingDirectory = CommandItem.StartIn Else WorkingDirectory = String.Empty

                                Admin = CommandItem.Admin

                                WindowProcessingStyle = CommandItem.StartingWindowState.ConvertStartingWindowStateToAProcessWindowStyle

                                If CommandItem.KeysToSend.Contains("$[") Then

                                    Dim ReplacementString As String = String.Empty
                                    ReplacementString = CommandItem.KeysToSend.Remove(0, CommandItem.KeysToSend.IndexOf("$[") + 2)
                                    ReplacementString = ReplacementString.Remove(ReplacementString.IndexOf("]"))
                                    KeysToSend = CommandItem.KeysToSend.Replace("$", VariablePortionOfIncomingMessage.Replace(" ", ReplacementString))
                                    KeysToSend = KeysToSend.Replace("[" & ReplacementString & "]", "")

                                Else

                                    If CommandItem.KeysToSend.Length > 0 Then KeysToSend = CommandItem.KeysToSend.Replace("$", VariablePortionOfIncomingMessage) Else KeysToSend = String.Empty

                                End If

                                If CommandItem.DesiredStatus = StatusValues.SwitchOn Then

                                    AnActionWasAttempted = True

                                    '*************************************************************************************************************

                                    If MakeItHappen Then
                                        ReturnValue = RunProgram(ProgramToRun, WorkingDirectory, Parameters, Admin, WindowProcessingStyle, KeysToSend)
                                    Else
                                        If IgnorSeparatingWords Then
                                            RedoRequired = True
                                            ReturnValue = ActionStatus.Unknown
                                            GoTo EarlyOut
                                        Else
                                            ReturnValue = RunProgram(ProgramToRun, WorkingDirectory, Parameters, Admin, WindowProcessingStyle, KeysToSend)
                                        End If
                                    End If

                                    Exit For

                                    '*************************************************************************************************************

                                Else

                                    FoundMatchingEntryButItWasSwitchedOff = True

                                    If IgnorSeparatingWords Then
                                        RedoRequired = True
                                    End If

                                End If

                            End If

                        End If

                    End If

                Next

            Next

            If AnActionWasAttempted Then

                If SpecialProcessing_NoMatchingPhrases Then
                    Log("command processed based on the 'No matching phrases' card")
                    ' Log("command processed based on the 'No matchingdoub phrases' card") ' removed after v4.8.3
                ElseIf SpecialProcessing_NoMatchingEnabledPhrases Then
                    Log("command processed based on the 'No matching enabled phrases' card")
                End If

            Else

                Dim WhatHappened As ActionStatus = ActionStatus.Unknown

                If SpecialProcessing_NoMatchingPhrases OrElse SpecialProcessing_NoMatchingEnabledPhrases Then

                Else

                    If FoundMatchingEntryButItWasSwitchedOff Then

                        WhatHappened = ActionIncomingIndivitualMessageNow("no matching enabled phrases " & IncomingMessage, RedoRequired, MakeItHappen) ' use recursion to check for special card 'added v3.4
                        ReturnValue = ActionStatus.NotProcessedWhileAtLeastOneMatchingPhraseWasFoundNoneWereEnabled

                    Else

                        WhatHappened = ActionIncomingIndivitualMessageNow("no matching phrases " & IncomingMessage, RedoRequired, MakeItHappen) ' use recursion to check for special card 'added v3.4
                        ReturnValue = ActionStatus.NotProcessedAsNoMatchingPhrasesFound

                    End If

                End If

            End If

        Catch ex As Exception

            Log(ex.ToString) 'todo: comment this out

        End Try

        ' used for debugging

        If gProduceDetailDumpWhenNoMatchIsFound Then

            Try

                If ReturnValue = ActionStatus.NotProcessedAsNoMatchingPhrasesFound Then

                    Log("++++++++++++++++++")

                    Log("Incoming message: *" & IncomingMessage & "*")

                    For Each CommandItem As gCommandRecord In CommandTable

                        With CommandItem

                            Log("******************")
                            Log(" ID: *" & .ID & "*")
                            Log(" ListenFor:")
                            Dim x As Integer = 1
                            Dim AllPhrasesToListenFor() As String = .ListenFor.Split(vbCr)
                            For Each SpecificPhraseToListenFor As String In AllPhrasesToListenFor
                                Log(x & ": *" & SpecificPhraseToListenFor.Trim & "*")
                                x += 1
                            Next

                        End With

                    Next

                    Log("------------------")

                End If

            Catch ex As Exception

            End Try

        End If

EarlyOut:

        Return ReturnValue

    End Function


    Friend Sub SpecialSubstitutions(ByVal ProgramTarget As String, ByVal ParmsTarget As String, ByVal VariablePortionOfIncomingMessage As String, ByRef ProgramToRun As String, ByRef Parameters As String)

        'v4.4 allow [Everything $] and [Everything $ .extention] 

        If ProgramTarget.ToLower.Contains("[everything $") Then

            Dim EverythingStartsAt As Integer = ProgramTarget.ToLower.IndexOf("[everything $")
            Dim EverythingEndsAt As Integer = ProgramTarget.ToLower.IndexOf("]", EverythingStartsAt)

            If EverythingEndsAt > -1 Then

                Dim SearchCriteria = ProgramTarget.Remove(0, EverythingStartsAt + 13)
                SearchCriteria = SearchCriteria.Remove(SearchCriteria.IndexOf("]"))
                SearchCriteria = VariablePortionOfIncomingMessage & " " & SearchCriteria

                Dim EverythingResult = EverythingSearch(SearchCriteria)


                ProgramTarget = ProgramTarget.Remove(EverythingStartsAt, EverythingEndsAt - EverythingStartsAt + 1)
                ProgramTarget = ProgramTarget.Insert(EverythingStartsAt, EverythingResult)

                ProgramToRun = ProgramTarget

            End If

        End If
        '' left off here

        If ParmsTarget.ToLower.Contains("[everything $") Then

            Dim EverythingStartsAt As Integer = ParmsTarget.ToLower.IndexOf("[everything $")
            Dim EverythingEndsAt As Integer = ParmsTarget.ToLower.IndexOf("]", EverythingStartsAt)

            If EverythingEndsAt > -1 Then

                Dim SearchCriteria = ParmsTarget.Remove(0, EverythingStartsAt + 13)
                SearchCriteria = SearchCriteria.Remove(SearchCriteria.IndexOf("]"))
                SearchCriteria = VariablePortionOfIncomingMessage & " " & SearchCriteria

                Dim EverythingResult = EverythingSearch(SearchCriteria)

                ParmsTarget = ParmsTarget.Remove(EverythingStartsAt, EverythingEndsAt - EverythingStartsAt + 1)
                ParmsTarget = ParmsTarget.Insert(EverythingStartsAt, EverythingResult)

                Parameters = ParmsTarget

            End If

        End If

        '***

        'v1.5  allow $[,]

        If ProgramTarget.Contains("$[") Then

            Dim ReplacementString As String
            ReplacementString = ProgramTarget.Remove(0, ProgramTarget.IndexOf("$[") + 2)
            ReplacementString = ReplacementString.Remove(ReplacementString.Length - 1)
            ProgramToRun = ProgramTarget.Replace("$", VariablePortionOfIncomingMessage.Replace(" ", ReplacementString))
            ProgramToRun = ProgramToRun.Replace("[" & ReplacementString & "]", "")

        Else

            ProgramToRun = ProgramTarget.Replace("$", VariablePortionOfIncomingMessage)

        End If

        If ParmsTarget.Contains("$[") Then

            Dim ReplacementString As String
            ReplacementString = ParmsTarget.Remove(0, ProgramTarget.IndexOf("$[") + 2)
            ReplacementString = ReplacementString.Remove(ReplacementString.IndexOf("]"))
            Parameters = ParmsTarget.Replace("$", VariablePortionOfIncomingMessage.Replace(" ", ReplacementString))
            Parameters = Parameters.Replace("[" & ReplacementString & "]", "")

        Else

            If ParmsTarget.Length > 0 Then Parameters = ParmsTarget.Replace("$", VariablePortionOfIncomingMessage) Else Parameters = String.Empty

        End If

        '***

    End Sub
    Private Sub ListView1_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles ListView1.MouseLeftButtonDown

        'deselects any selected rows, preventing the double click on an empty space to be applied to the previousily selected rows
        ListView1.SelectedIndex = -1

    End Sub

#Region "Status Bar"

    'Enum PushServiceStatus
    '    Undefined = 0
    '    Good = 1
    '    Warning = 2
    '    bad = 3
    'End Enum

    Private MaxNumberOfSecondsAProcessMessageWillStayInTheTaskBar As Integer

    Private GoodConnectionColour As New SolidColorBrush()
    Private BadConnectionColour As New SolidColorBrush()

    Private ActionSuccessfulResultBrush As Brush

    Private ActionUnknowBrush As Brush
    Private ActionDidNotRunAsExpected As Brush
    Private ActionUnknownResultBrush As Brush
    Private ActionFailedResultBrush As Brush
    Private ActionPartiallySucceededResultBrush As Brush

    Private Const GoodConnectionThreshold As Integer = 65

    Private Sub SetupStatusBar()

        MaxNumberOfSecondsAProcessMessageWillStayInTheTaskBar = 5

        GoodConnectionColour.Color = Colors.Black
        BadConnectionColour.Color = Colors.Red

        ActionSuccessfulResultBrush = Brushes.DarkGreen
        ActionPartiallySucceededResultBrush = Brushes.LightGoldenrodYellow
        ActionDidNotRunAsExpected = Brushes.DarkOrange
        ActionUnknownResultBrush = Brushes.Black
        ActionFailedResultBrush = Brushes.Red

        ConfirmMQTTConnectionTimer.Interval = CheckForGoodConnectionEverySecond
        ConfirmMQTTConnectionTimer.Start()

        ConfirmPushbulletConnectionTimer.Interval = CheckForGoodConnectionEverySecond
        ConfirmPushbulletConnectionTimer.Start()

        ConfirmPushoverConnectionTimer.Interval = CheckForGoodConnectionEverySecond
        ConfirmPushoverConnectionTimer.Start()

    End Sub

    Private Sub UpdateTheStatusBar(ByVal Input As String, Optional ByVal WhatHappened As ActionStatus = ActionStatus.Unknown)

        If Me.Visibility = Visibility.Visible Then
            Me.Dispatcher.Invoke(New UpdateStatusBarCallback(AddressOf UpdateStatusBar), New Object() {Input, WhatHappened})
        End If

    End Sub

    Delegate Sub UpdateStatusBarCallback(ByVal message As String, ByVal WhatHappened As Integer)  ' WhatHappened should be of type ActionStatus, however this would cause a compiler error so integer is used instead
    Private Sub UpdateStatusBar(ByVal message As String, ByVal WhatHappened As ActionStatus)

        If message.Length > 0 Then

            ProcessingStatus.Content = message

            Select Case WhatHappened

                Case Is = ActionStatus.Succeeded
                    ProcessingStatus.Foreground = ActionSuccessfulResultBrush

                Case Is = ActionStatus.PartiallySucceeded
                    ProcessingStatus.Foreground = ActionPartiallySucceededResultBrush

                Case Is = ActionStatus.MasterSwitchWasOff, ActionStatus.NotProcessedAsNoMatchingPhrasesFound, ActionStatus.NotProcecessAsAUACPromptWouldBeRequired
                    ProcessingStatus.Foreground = ActionDidNotRunAsExpected

                Case Is = ActionStatus.Failed
                    ProcessingStatus.Foreground = ActionFailedResultBrush

                Case Is = ActionStatus.Unknown
                    ProcessingStatus.Foreground = ActionUnknownResultBrush

            End Select

        End If

    End Sub

    Dim WithEvents UndoDisplayTimer As System.Windows.Forms.Timer = New System.Windows.Forms.Timer

    Private Sub UpdateUndoDisplayMessage(ByVal Message As String)

        Me.Dispatcher.Invoke(New Force_UpdateUndoDisplayMessageCallback(AddressOf Force_UpdateUndoDisplayMessageNow), New Object() {Message})

    End Sub

    Delegate Sub Force_UpdateUndoDisplayMessageCallback(ByVal Message As String)
    Private Sub Force_UpdateUndoDisplayMessageNow(ByVal Message As String)

        UpdateTheStatusBar(Message, ActionStatus.LeaveBlank)

    End Sub

    Private Sub UndoDisplayTimer_tick(sender As Object, e As EventArgs) Handles UndoDisplayTimer.Tick

        Static TimeAlive As Integer = 0

        TimeAlive += 1

        If TimeAlive > 10 Then
            UpdateUndoDisplayMessage(" ")
            UndoDisplayTimer.Stop()
            TimeAlive = 0
        End If

    End Sub

    '*****************************************************************************************************************************************************************

    Dim WithEvents ConfirmMQTTConnectionTimer As System.Windows.Forms.Timer = New System.Windows.Forms.Timer

    'Delegate Sub Force_ResetTheMQTTDisplayTimerCallback(ByVal RevisedLastTimeDataWasReceivedFromMQTT As Date)
    'Private Sub Force_ResetTheMQTTDisplayTimerNow(ByVal RevisedLastTimeDataWasReceivedFromMQTT As Date)
    'End Sub

    Private Sub ConfirmMQTTConnectionTimer_Tick(sender As Object, e As EventArgs) Handles ConfirmMQTTConnectionTimer.Tick

        If My.Settings.UseMQTT Then

            If (gMQTTConnectionStatus = ConnectionStatus.Connected) OrElse (gMQTTConnectionStatus = ConnectionStatus.Unknown) Then

                If ConfirmMQTTConnectionTimer.Interval = CheckForGoodConnectionEvery15Seconds Then ConfirmMQTTConnectionTimer.Interval = CheckForGoodConnectionEverySecond

            Else

                ConfirmMQTTConnectionTimer.Interval = CheckForGoodConnectionEvery15Seconds

                ' try to reconnect
                Log("Attempting to reconnect to the MQTT server")
                GetMQTTUnderway()

                If (gMQTTConnectionStatus = ConnectionStatus.Connected) Then
                    ConfirmMQTTConnectionTimer.Interval = CheckForGoodConnectionEverySecond
                Else
                    ConfirmMQTTConnectionTimer.Interval = CheckForGoodConnectionEvery15Seconds
                End If

            End If

        End If

    End Sub

    '*****************************************************************************************************************************************************************

    Dim WithEvents ConfirmPushbulletConnectionTimer As System.Windows.Forms.Timer = New System.Windows.Forms.Timer

    Private Sub ResetTheTimePushbulletWasLastAccessed(ByVal RevisedLastTimeDataWasReceivedFromPushbullet As Date)

        Me.Dispatcher.Invoke(New Force_ResetThePushbulletDisplayTimerCallback(AddressOf Force_ResetPushbulletTheDisplayTimerNow), New Object() {RevisedLastTimeDataWasReceivedFromPushbullet})

    End Sub

    Delegate Sub Force_ResetThePushbulletDisplayTimerCallback(ByVal RevisedLastTimeDataWasReceivedFromPushbullet As Date)
    Private Sub Force_ResetPushbulletTheDisplayTimerNow(ByVal RevisedLastTimeDataWasReceivedFromPushbullet As Date)

        LastTimeDataWasReceivedFromPushbullet = RevisedLastTimeDataWasReceivedFromPushbullet  ' this forces the ConfirmPushbulletConnectionTimer_Tick to re-evaluate the Pushbullet status

        If LastTimeDataWasReceivedFromPushbullet > Now.AddSeconds(-GoodConnectionThreshold) Then
            If ConfirmPushbulletConnectionTimer.Interval = CheckForGoodConnectionEvery15Seconds Then ConfirmPushbulletConnectionTimer.Interval = CheckForGoodConnectionEverySecond
        End If

    End Sub

    '*****************************************************************************************************************************************************************
    Dim WithEvents ConfirmPushoverConnectionTimer As System.Windows.Forms.Timer = New System.Windows.Forms.Timer

    Private Sub ResetTheTimePushoverWasLastAccessed(ByVal RevisedLastTimeDataWasReceivedFromPushover As Date)

        Me.Dispatcher.Invoke(New Force_ResetThePushoverDisplayTimerCallback(AddressOf Force_ResetThePushoverDisplayTimerNow), New Object() {RevisedLastTimeDataWasReceivedFromPushover})

    End Sub

    Delegate Sub Force_ResetThePushoverDisplayTimerCallback(ByVal RevisedLastTimeDataWasReceivedFromPushover As Date)
    Private Sub Force_ResetThePushoverDisplayTimerNow(ByVal RevisedLastTimeDataWasReceivedFromPushover As Date)

        LastTimeDataWasReceivedFromPushover = RevisedLastTimeDataWasReceivedFromPushover  ' this forces the ConfirmPushoverConnectionTimer_Tick to re-evaluate the Pushover status

        If LastTimeDataWasReceivedFromPushover > Now.AddSeconds(-GoodConnectionThreshold) Then
            If ConfirmPushoverConnectionTimer.Interval = CheckForGoodConnectionEvery15Seconds Then ConfirmPushoverConnectionTimer.Interval = CheckForGoodConnectionEverySecond
        End If

    End Sub

    Private Sub ConfirmPushbulletConnectionTimer_Tick(sender As Object, e As EventArgs) Handles ConfirmPushbulletConnectionTimer.Tick

        Try

            UpdateStatusBar()

            If Not My.Settings.UsePushbullet Then
                Exit Sub
            End If

        Catch ex As Exception

        End Try

        Static Dim NumberOfSecondsLastMessageHasBeenOnScreen As Integer = 0
        Static Dim LastMessageRecieved As String = String.Empty
        Static Dim DoOnce As Boolean = True

        Const MissingAPIKeyMessage As String = "The Pushbullet Access Token (Actions - Options - Pushbullet Access Token) is missing"
        Const NoNetworkMessage As String = "Your computer appears not to be connected to your network"
        Const TestModeMessage As String = "*** TEST MODE ***"

        If PushbulletEnbledInTesting Then
        Else
            If DoOnce Then
                DoOnce = False
            Else
                Exit Sub
            End If
        End If

        If My.Settings.UsePushbullet Then

            If LastTimeDataWasReceivedFromPushbullet > Now.AddSeconds(-GoodConnectionThreshold) Then

                gPushbulletConnectionStatus = ConnectionStatus.Connected

                PushBulletCommunicationsStatus.Foreground = GoodConnectionColour
                PushBulletCommunicationsStatus.Content = "Pushbullet connected"

                If ProcessingStatus.Content = NoNetworkMessage Then ProcessingStatus.Content = String.Empty

                If ConfirmPushbulletConnectionTimer.Interval = CheckForGoodConnectionEvery15Seconds Then ConfirmPushbulletConnectionTimer.Interval = CheckForGoodConnectionEverySecond

            Else

                gPushbulletConnectionStatus = ConnectionStatus.DisconnectedorClosed

                PushBulletCommunicationsStatus.Foreground = BadConnectionColour
                PushBulletCommunicationsStatus.Content = "Pushbullet not connected"

                ProcessingStatus.Foreground = ActionFailedResultBrush

                If My.Settings.PushBulletAPI = String.Empty Then
                    ProcessingStatus.Content = MissingAPIKeyMessage

                ElseIf CurrentNetworkStatus <> OperationalStatus.Up Then
                    ProcessingStatus.Content = NoNetworkMessage

                ElseIf PushbulletEnbledInTesting Then
                    ProcessingStatus.Content = String.Empty

                Else
                    ProcessingStatus.Content = TestModeMessage

                End If

                If PushbulletEnbledInTesting Then

                    If CurrentNetworkStatus = OperationalStatus.Up Then

                        ConfirmPushbulletConnectionTimer.Interval = CheckForGoodConnectionEvery15Seconds
                        Log("Attempting to reconnect to Pushbullet")
                        WebSocketErrorWasReportedPushbullet = OpenThePushbulletWebSocket()

                    End If

                End If

            End If

        End If


        If ProcessingStatus.Content <> String.Empty Then

            'keep these messages in place: MissingAPIKeyMessage, NoNetworkMessage, TestModeMessage
            'otherwise clear the processing status line as needed

            If (gPushbulletConnectionStatus = ConnectionStatus.DisconnectedorClosed) AndAlso (ProcessingStatus.Content = MissingAPIKeyMessage) OrElse (ProcessingStatus.Content = NoNetworkMessage) OrElse (ProcessingStatus.Content = TestModeMessage) Then

            Else

                If ProcessingStatus.Content = LastMessageRecieved Then
                    If NumberOfSecondsLastMessageHasBeenOnScreen >= MaxNumberOfSecondsAProcessMessageWillStayInTheTaskBar Then
                        NumberOfSecondsLastMessageHasBeenOnScreen = 0

                        If PushbulletEnbledInTesting Then
                            ProcessingStatus.Content = String.Empty
                        Else
                            ProcessingStatus.Content = TestModeMessage
                            ProcessingStatus.Foreground = ActionFailedResultBrush
                        End If

                    Else
                        NumberOfSecondsLastMessageHasBeenOnScreen += 1
                    End If
                Else
                    LastMessageRecieved = ProcessingStatus.Content
                    NumberOfSecondsLastMessageHasBeenOnScreen = 0
                End If

            End If

        End If

        ConfirmSystrayIcons()

    End Sub

    '*****************************************************************************************************************************************************************

    Delegate Sub UpdateStatusBarNowCallback()
    Private Sub UpdateStatusBar()

        If (Not My.Settings.UseDropbox) AndAlso (Not My.Settings.UsePushbullet) AndAlso (Not My.Settings.UsePushover) AndAlso (Not My.Settings.UseMQTT) AndAlso (Not My.Settings.SuppressStartupNotice) Then

            StatusLine1.Visibility = Visibility.Visible
            PushBulletCommunicationsStatus.Visibility = Visibility.Visible
            PushBulletCommunicationsStatus.Content = "Usually at least one of Dropbox, Pushbullet, Pushover, or MQTTT should be enabled for use.  Currently none are."
            PushBulletCommunicationsStatus.Foreground = BadConnectionColour

            StatusLine2.Visibility = Visibility.Collapsed
            PushoverCommunicationsStatus.Visibility = Visibility.Collapsed

            StatusLine3.Visibility = Visibility.Collapsed
            MQTTCommunicationsStatus.Visibility = Visibility.Collapsed

            ConfirmSystrayIcons()

            Exit Sub

        End If

        ' *****************************

        If My.Settings.UsePushbullet Then
        Else
            PushBulletCommunicationsStatus.Content = String.Empty
        End If

        If My.Settings.UsePushover Then
        Else
            PushoverCommunicationsStatus.Content = String.Empty
        End If

        If My.Settings.UseMQTT Then
            MQTTCommunicationsStatus.Content = gMQTTStatusText
            If MQTTCommunicationsStatus.Content = "MQTT connected" Then
                MQTTCommunicationsStatus.Foreground = GoodConnectionColour
            Else
                MQTTCommunicationsStatus.Foreground = BadConnectionColour
            End If
        Else
            MQTTCommunicationsStatus.Content = String.Empty
        End If

        StatusLine1.Visibility = Visibility.Visible = My.Settings.UsePushbullet
        StatusLine2.Visibility = Visibility.Visible = My.Settings.UsePushover
        StatusLine3.Visibility = Visibility.Visible = My.Settings.UseMQTT

        IIf(My.Settings.UsePushbullet, PushBulletCommunicationsStatus.Visibility = Visibility.Visible, PushBulletCommunicationsStatus.Visibility = Visibility.Collapsed)
        IIf(My.Settings.UsePushover, PushoverCommunicationsStatus.Visibility = Visibility.Visible, PushoverCommunicationsStatus.Visibility = Visibility.Collapsed)
        IIf(My.Settings.UseMQTT, MQTTCommunicationsStatus.Visibility = Visibility.Visible, MQTTCommunicationsStatus.Visibility = Visibility.Collapsed)

    End Sub

    Private Sub ConfirmPushoverConnectionTimer_Tick(sender As Object, e As EventArgs) Handles ConfirmPushoverConnectionTimer.Tick

        Static Dim NumberOfSecondsLastMessageHasBeenOnScreen As Integer = 0
        Static Dim LastMessageRecieved As String = String.Empty
        Static Dim DoOnce As Boolean = True

        Const ProblemWithUserIDPasswordMessage As String = "Pushover authentication failed"
        Const NoNetworkMessage As String = "Your computer appears not to be connected to your network"
        Const TestModeMessage As String = "*** TEST MODE ***"

        If PushoverEnabledInTesting Then
        Else
            If DoOnce Then
                DoOnce = False
            Else
                Exit Sub
            End If
        End If

        If My.Settings.UsePushover Then

            If LastTimeDataWasReceivedFromPushover > Now.AddSeconds(-GoodConnectionThreshold) Then

                gPushoverConnectionStatus = ConnectionStatus.Connected

                PushoverCommunicationsStatus.Foreground = GoodConnectionColour
                PushoverCommunicationsStatus.Content = "Pushover connected"

                If ProcessingStatus.Content = NoNetworkMessage Then ProcessingStatus.Content = String.Empty

                If ConfirmPushoverConnectionTimer.Interval = CheckForGoodConnectionEvery15Seconds Then ConfirmPushoverConnectionTimer.Interval = CheckForGoodConnectionEverySecond

            Else

                gPushoverConnectionStatus = ConnectionStatus.DisconnectedorClosed

                PushoverCommunicationsStatus.Foreground = BadConnectionColour
                PushoverCommunicationsStatus.Content = "Pushover not connected"

                ProcessingStatus.Foreground = ActionFailedResultBrush
                If Not ArePushoverIDAndSecretAvailable() Then
                    ProcessingStatus.Content = ProblemWithUserIDPasswordMessage

                ElseIf CurrentNetworkStatus <> OperationalStatus.Up Then
                    ProcessingStatus.Content = NoNetworkMessage

                ElseIf PushoverEnabledInTesting Then
                    ProcessingStatus.Content = String.Empty

                Else
                    ProcessingStatus.Content = TestModeMessage

                End If

                If PushoverEnabledInTesting Then

                    If CurrentNetworkStatus = OperationalStatus.Up Then

                        ConfirmPushoverConnectionTimer.Interval = CheckForGoodConnectionEvery15Seconds
                        If gCriticalPushOverErrorReported Then
                        Else
                            Log("Attempting to reconnect to Pushover")
                            OpenThePushoverWebSocket()
                        End If

                    End If

                End If

            End If

        End If

        If ProcessingStatus.Content <> String.Empty Then

            'keep error messages in place
            'otherwise clear the processing status line as needed

            If (gPushoverConnectionStatus = ConnectionStatus.DisconnectedorClosed) AndAlso (ProcessingStatus.Content = ProblemWithUserIDPasswordMessage) OrElse (ProcessingStatus.Content = NoNetworkMessage) OrElse (ProcessingStatus.Content = TestModeMessage) Then

            Else

                If ProcessingStatus.Content = LastMessageRecieved Then
                    If NumberOfSecondsLastMessageHasBeenOnScreen >= MaxNumberOfSecondsAProcessMessageWillStayInTheTaskBar Then
                        NumberOfSecondsLastMessageHasBeenOnScreen = 0

                        If PushoverEnabledInTesting Then
                            ProcessingStatus.Content = String.Empty
                        Else
                            ProcessingStatus.Content = TestModeMessage
                            ProcessingStatus.Foreground = ActionFailedResultBrush
                        End If

                    Else
                        NumberOfSecondsLastMessageHasBeenOnScreen += 1
                    End If
                Else
                    LastMessageRecieved = ProcessingStatus.Content
                    NumberOfSecondsLastMessageHasBeenOnScreen = 0
                End If

            End If

        End If

        ConfirmSystrayIcons()

    End Sub


#Region "Drag and Drop"

    Private Sub ListView1_Drop(sender As Object, e As System.Windows.DragEventArgs) ' Handles ListView1.Drop (handle not required as the xaml calls this routine)

        If e.Data.GetDataPresent("Shell IDList Array") Then

            ' Dropping a short cut

            ' get the Sort Order of row under mouse

            ' to get the Sort of the row under the mouse we first need to find the row
            ' if the client drops the icon on a textbox the sort order can be determined right away
            ' however if the client does not drop the icon on a text box the sort order needs to be determined by finding the closest row
            ' Testing starts in x offset 100 where there should be a text box of some kind
            ' for the y offset, testing starts a the y offset where the mouse is at currently and then goes up and down by 1 point until hopefully eventually a hit is achieved
            ' testing does not go up above or down below where the listview starts and ends respectively

            Dim SortOrderOfRowUnderMouse As Integer = 10
            Dim pt As System.Windows.Point
            Dim obj As Object

            Try

                pt = e.GetPosition(Me)
                obj = System.Windows.Media.VisualTreeHelper.HitTest(Me, pt)
                If obj.visualhit.ToString = "System.Windows.Controls.TextBlock" Then
                    SortOrderOfRowUnderMouse = obj.VisualHit.Parent.content.SortOrder
                    Exit Try
                End If

                Dim YStartingOffset As Integer = pt.Y
                Dim YLimitToSearch As Integer = YStartingOffset + ListView1.ActualHeight
                Dim Increment As Integer = 0

                For y As Integer = YStartingOffset To YLimitToSearch

                    Increment += 1

                    'test going downward for as far as the list view extends
                    If (YStartingOffset + Increment) <= ListView1.ActualHeight Then
                        obj = System.Windows.Media.VisualTreeHelper.HitTest(Me, New Point(100, YStartingOffset + Increment))
                        If obj.visualhit.ToString = "System.Windows.Controls.TextBlock" Then
                            SortOrderOfRowUnderMouse = obj.VisualHit.Parent.content.SortOrder
                            Exit Try
                        End If
                    End If

                    'test going upward until the top of the list view
                    If (YStartingOffset - Increment) > 0 Then
                        obj = System.Windows.Media.VisualTreeHelper.HitTest(Me, New Point(100, YStartingOffset - Increment))
                        If obj.visualhit.ToString = "System.Windows.Controls.TextBlock" Then
                            SortOrderOfRowUnderMouse = obj.VisualHit.Parent.content.SortOrder
                            Exit Try
                        End If
                    End If

                Next

            Catch ex As Exception
            End Try

            For x = 0 To ListView1.Items.Count - 1
                If ListView1.Items.Item(x).SortOrder = SortOrderOfRowUnderMouse Then
                    ListView1.SelectedItem = ListView1.Items.Item(x)
                    Exit For 'v1.6 added exit for
                End If
            Next

            obj = Nothing

            AddgCurrentlySelectedRecordIntoTheDatabase(True, "", SortOrderOfRowUnderMouse, e)

        Else

            ' Dropping some text from the Session Log to update Listen for

            Try

                'Get the incoming data

                Dim IncomingData As String = e.Data.GetData(System.Windows.DataFormats.Text)

                DragAndDropUnderway = False

                e.Handled = True

                If (IncomingData Is Nothing) OrElse (IncomingData.Length = 0) Then Exit Try  ' v1.7

                'close and reload the session log window so that another drag and drop can be performed
                MenuViewSessionLog.IsChecked = True 'v1.4
                ToggleTheSessionLogOpenAndClosed()

                'Get the data associated with the row to be updated
                Dim pt As System.Windows.Point = e.GetPosition(Me)
                Dim obj As Object = System.Windows.Media.VisualTreeHelper.HitTest(Me, pt)

                Dim TextBlock As TextBlock = TryCast(obj.visualhit, TextBlock)
                Dim Parent As Object = TextBlock.Parent

                Dim ID As Integer = Parent.Content.ID
                Dim SortOrder As Integer = Parent.Content.SortOrder
                Dim ExistingListenForString As String = Parent.Content.ListenFor


                'Update the existing ListenFor string to include the incoming data
                Dim UpdatedListenForString As String = ExistingListenForString & vbCrLf & IncomingData
                UpdatedListenForString = StringSort(UpdatedListenForString)

                If ExistingListenForString = UpdatedListenForString Then
                Else

                    'Update the database

                    AddToUndoTable(UndoRational.drop)

                    Dim DatabaseRow As MyTable1Class

                    gBossLoadUnderway = True

                    DatabaseRow = ReadARecord(ID)
                    DatabaseRow.ListenFor = UpdatedListenForString

                    With DatabaseRow
                        ChangeARecord(.ID, .SortOrder, .DesiredStatus, .WorkingStatus, .Description, .ListenFor, .Open, .Parameters, .StartIn, .Admin, .StartingWindowState, .KeysToSend)
                    End With

                    RefreshListView(gCurrentlySelectedRow.SortOrder)

                    LoadFromDatabase()

                    DatabaseRow = Nothing

                    gBossLoadUnderway = False

                End If

            Catch ex As Exception

            End Try

        End If

    End Sub

    Friend Sub SafelyAddgCurrentlySelectedRecordIntoTheDatabase(ByVal LoadViaDragAndDrop As Boolean, ByVal FileName As String)
        Dispatcher.BeginInvoke(Sub() AddgCurrentlySelectedRecordIntoTheDatabase(LoadViaDragAndDrop, FileName))
    End Sub

    Private Sub AddgCurrentlySelectedRecordIntoTheDatabase(ByVal LoadViaDragAndDrop As Boolean, ByVal FileName As String, Optional ByVal SortOrderOfRowUnderMouse As Integer = 10, Optional ByRef e As System.Windows.DragEventArgs = Nothing)

        Dim DroppedFileName As String = String.Empty

        If e Is Nothing Then
            ' .p2r file was double clicked
            DroppedFileName = FileName

        Else
            ' .p2r file was dragged and dropped
            DroppedFileName = e.Data.GetData(System.Windows.DataFormats.FileDrop)(0)

        End If

        If (DetermineP2RFileType(DroppedFileName) = P2RFileType.ExportStyleMayContainMultipleCards) Then

            gImportFileName = DroppedFileName
            DoImport()

        Else

            Try

                With gCurrentlySelectedRow

                    Dim OpenUpAddEditWindow As Boolean = True

                    If LoadViaDragAndDrop Then

                        .ID = GetMaxIDFromDatabase() + 1
                        .SortOrder = SortOrderOfRowUnderMouse + 5
                        .DesiredStatus = StatusValues.SwitchOff
                        .WorkingStatus = StatusValues.SwitchOff
                        .Description = String.Empty
                        .ListenFor = String.Empty
                        .Open = String.Empty
                        .StartIn = String.Empty
                        .Parameters = String.Empty
                        .Admin = False
                        .KeysToSend = String.Empty

                        OpenUpAddEditWindow = DropIntoCurrentlySelectedRow(e)

                    End If

                    If OpenUpAddEditWindow Then
                    Else
                        ' this will happen when the file that was dragged and dropped was as an exported file
                        RefreshListView(gCurrentlySelectedRow.SortOrder)
                        LoadFromDatabase()
                        Exit Sub
                    End If

                    InsertARecord(.SortOrder, StatusValues.SwitchOff, StatusValues.SwitchOff, .Description, .ListenFor, .Open, .Parameters, .StartIn, .Admin, .StartingWindowState, .KeysToSend)

                    Dim CurrentRecord As MyTable1Class = ReadARecord(GetMaxIDFromDatabase)

                    With gCurrentlySelectedRow

                        .ID = CurrentRecord.ID
                        .SortOrder = CurrentRecord.SortOrder
                        .DesiredStatus = CurrentRecord.DesiredStatus
                        .WorkingStatus = CurrentRecord.WorkingStatus
                        .Description = CurrentRecord.Description
                        .ListenFor = CurrentRecord.ListenFor
                        .Open = CurrentRecord.Open
                        .Parameters = CurrentRecord.Parameters
                        .StartIn = CurrentRecord.StartIn
                        .Admin = CurrentRecord.Admin
                        .StartingWindowState = CurrentRecord.StartingWindowState
                        .KeysToSend = CurrentRecord.KeysToSend

                    End With

                    gDropInProgress = True

                    PerformAction("Edit")

                    gDropInProgress = False

                    If EditWasCancelled Then
                        DeleteARecord(.ID)
                    End If

                    If MenuSort.IsChecked Then
                        RebuildTable1(True)
                    End If

                    RefreshListView(gCurrentlySelectedRow.SortOrder)

                    LoadFromDatabase()

                End With

            Catch ex As Exception

            End Try

        End If

    End Sub


    Friend Sub RebuildTable1(ByVal SortByDescription As Boolean)

        Try

            RebuildTable1_common(SortByDescription)

        Catch ex As Exception
        End Try

    End Sub

    Private Sub ListView1_DragEnter(sender As Object, e As System.Windows.DragEventArgs) Handles ListView1.DragEnter

        If (e.Data.GetDataPresent(System.Windows.DataFormats.Text)) Then
            e.Effects = System.Windows.DragDropEffects.Copy
        Else
            e.Effects = System.Windows.DragDropEffects.None
        End If

    End Sub

    Private Sub ListView1_DragLeave(sender As Object, e As System.Windows.DragEventArgs) Handles ListView1.DragLeave

        e.Effects = System.Windows.DragDropEffects.None

    End Sub

    Private Sub WindowBoss_ContentRendered(sender As Object, e As EventArgs) Handles Me.ContentRendered
        Keyboard.Focus(ListView1)
    End Sub

    Private Sub ThisWindow_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged

        KeepHelpOnTop()

    End Sub


#End Region

#Region "Update settings as necessary"


    Private Sub UpdateSettingsAsNecissary()

        Try

            Dim SettingsNeedToBeSave As Boolean = False

            'Upgrade and carry forward settings as necessary

            Dim a As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
            Dim appVersion As Version = a.GetName().Version
            Dim appVersionString As String = appVersion.ToString

            If My.Settings.ApplicationVersion <> appVersion.ToString Then

                UgradeToStronglyNamedSettings()  'v4.3

                ' v4.6 - starting in v4.6 this works differently, hence MySettingsKludgeUpdate() is replaces My.Settings.Upgrade() 
                ' My.Settings.Upgrade()  
                MySettingsKludgeUpdate()

                My.Settings.ApplicationVersion = appVersionString
                SettingsNeedToBeSave = True
            End If

            'validate next version check date 
            If (My.Settings.NextVersionCheckDate = Nothing) OrElse (Not IsDate(My.Settings.NextVersionCheckDate)) Then
                My.Settings.NextVersionCheckDate = Today.AddYears(-1)
                SettingsNeedToBeSave = True
            End If

            'validate unique id
            If My.Settings.MyUniqueID = String.Empty Then
                My.Settings.MyUniqueID = GenerateRandomPassword(16)
                SettingsNeedToBeSave = True
            End If

            'validate Pushbullet title filter
            If My.Settings.PushBulletTitleFilter.Length = 0 Then
                My.Settings.PushBulletTitleFilter = "Push2Run " & My.Computer.Name
                SettingsNeedToBeSave = True
            End If

            If SettingsNeedToBeSave Then
                My.Settings.Save()
            End If


        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try


    End Sub

    Private Sub MySettingsKludgeUpdate()

        Try

            Dim settingValuesFromPreviousSettingsFile As String = String.Empty

            Dim myProgram As String = My.Application.Info.AssemblyName

            Dim strongNameSettingsFolderForThisVersion As String = Path.GetDirectoryName(System.Configuration.ConfigurationManager.OpenExeConfiguration(System.Configuration.ConfigurationUserLevel.PerUserRoamingAndLocal).FilePath)

            If Directory.Exists(strongNameSettingsFolderForThisVersion) Then

                'this should not happen on an upgrade, so if this directory is found our job here is done 
                Exit Sub
            Else

                'the following commented line of code does not work as expected due to latency issues; see https://stackoverflow.com/questions/2225601/directory-createdirectory-latency-issue
                '  Directory.CreateDirectory(Path.GetDirectoryName(strongNameSettingsFolderForThisVersion))
                'using the following instead:

                Dim workingDir As New DirectoryInfo(strongNameSettingsFolderForThisVersion)
                workingDir.Create()
                workingDir.Refresh()

            End If

            ' Get setting values from previous settings file

            Dim parentOfStrongNameSettingsFolderForThisVersion As String = Path.GetDirectoryName(strongNameSettingsFolderForThisVersion)
            Dim dir As New DirectoryInfo(parentOfStrongNameSettingsFolderForThisVersion)


            If dir.GetDirectories.Count = 0 Then

                'if the parent of the StrongNamedSettingsFolder has no children then approach this as a first time install
                'upgrade can proceed without the Kludge

                My.Settings.Upgrade()
                My.Settings.Save()
                My.Settings.Reload()
                Exit Sub

            End If

            ' find the directory with the greatest previous version number 

            Dim folderWithGreatestPreviousVersionNumber As String = String.Empty

            For Each individualDirectory In dir.GetDirectories()

                If individualDirectory.FullName > folderWithGreatestPreviousVersionNumber Then

                    If individualDirectory.FullName <> strongNameSettingsFolderForThisVersion Then
                        folderWithGreatestPreviousVersionNumber = individualDirectory.FullName
                    End If

                End If

            Next

            Try

                settingValuesFromPreviousSettingsFile = File.ReadAllText(folderWithGreatestPreviousVersionNumber & "\user.config")

            Catch ex As Exception

                ' this is a first time install
                '  upgrade can proceed without the Kludge

                My.Settings.Upgrade()
                My.Settings.Save()
                My.Settings.Reload()
                Exit Sub

            End Try


            If settingValuesFromPreviousSettingsFile.Contains("System.Configuration.ClientSettingsSection") Then

                ' the previous settings file has already been adjusted so the current settings file will not need to be adjusted again, upgrade can proceed without the Kludge
                My.Settings.Upgrade()
                My.Settings.Save()
                My.Settings.Reload()
                Exit Sub

            End If

            ' Establish the header, footer, and core values for the new settings file

            Dim header = "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf &
                             "<configuration>" & vbCrLf &
                             "    <configSections>" & vbCrLf &
                             "        <sectionGroup name=""userSettings"">" & vbCrLf &
                             "            <section name=""" & myProgram & ".My.MySettings"" type=""System.Configuration.ClientSettingsSection, System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"" allowExeDefinition=""MachineToLocalUser"" requirePermission=""false""/>" & vbCrLf &
                             "        </sectionGroup>" & vbCrLf &
                             "    </configSections>" & vbCrLf &
                             "    <userSettings>" & vbCrLf &
                             "        <" & myProgram & ".My.MySettings>" & vbCrLf

            Dim footer = "        </" & myProgram & ".My.MySettings>" & vbCrLf &
                         "    </userSettings>" & vbCrLf &
                         "</configuration>"


            Dim coreSettingValuesFromPreviousSettingsFile As String = String.Empty

            If settingValuesFromPreviousSettingsFile > String.Empty Then
                coreSettingValuesFromPreviousSettingsFile = settingValuesFromPreviousSettingsFile.Remove(0, settingValuesFromPreviousSettingsFile.IndexOf("<setting name"))
                coreSettingValuesFromPreviousSettingsFile = coreSettingValuesFromPreviousSettingsFile.Remove(coreSettingValuesFromPreviousSettingsFile.LastIndexOf("</setting>") + "</setting>".Length())
            End If

            'do the Kludge upgrade
            File.WriteAllText(strongNameSettingsFolderForThisVersion & "\user.config", header & coreSettingValuesFromPreviousSettingsFile & footer)
            My.Settings.Reload()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub UgradeToStronglyNamedSettings()

        ' note this routine works on systems where the application has been deployed in either a release or a debug configuration on a machine, but not both

        Try

            Dim myProgram As String = My.Application.Info.AssemblyName

            Dim strongNameSettingsFolderForThisVersion As String = Path.GetDirectoryName(System.Configuration.ConfigurationManager.OpenExeConfiguration(System.Configuration.ConfigurationUserLevel.PerUserRoamingAndLocal).FilePath)

            'if the Strong Name settings folder for this version already exists then we have already converted a Non Strong Name settings folder if one had existed and our job here is done
            If Directory.Exists(strongNameSettingsFolderForThisVersion) Then Exit Sub

            Dim parentOfStrongNameSettingsFolderForThisVersion As String = Path.GetDirectoryName(strongNameSettingsFolderForThisVersion)

            'if the Parent of the Strong Name settings folder for this version already exists then we have already converted a Non Strong Name settings folder if one had existed and our job here is done
            If Directory.Exists(parentOfStrongNameSettingsFolderForThisVersion) Then Exit Sub

            Dim grandParentOfStrongNameSettingsFolderForThisVersion As String = strongNameSettingsFolderForThisVersion.Remove(strongNameSettingsFolderForThisVersion.IndexOf("\" & myProgram & ".exe_StrongName_"))

            ' in the grand parent of the strong name folder for this version there should be either:
            ' no children with a non strong name for this program
            ' or just one child with a non strong name for this program 

            ' if there are no children we are done
            ' if there is a child we need to copy it over to the new strong name folder so that it can be upgraded

            Dim parentOfNonStrongNameSettingsFolder As String = String.Empty

            Dim dir As New DirectoryInfo(grandParentOfStrongNameSettingsFolderForThisVersion)

            Try

                For Each individualDirectory In dir.GetDirectories()

                    If individualDirectory.FullName.Contains(myProgram & ".exe_Url_") Then

                        parentOfNonStrongNameSettingsFolder = individualDirectory.FullName
                        Exit For

                    End If

                Next

            Catch ex As Exception
                ' if an exception is thrown this is a first time install and our job here is done
                Exit Sub
            End Try

            ' if there were no children with non strong names we are done
            If parentOfNonStrongNameSettingsFolder = String.Empty Then Exit Sub

            ' there is a child so we need to copy it over to the new strong name folder in order that may be upgraded

            ' Find the non strong name directory with the greatest version number
            Dim nonStrongNameSettingsFolderWithGreatestVersionNumber As String = String.Empty

            dir = New DirectoryInfo(parentOfNonStrongNameSettingsFolder)

            For Each individualDirectory In dir.GetDirectories()

                If individualDirectory.FullName > nonStrongNameSettingsFolderWithGreatestVersionNumber Then

                    nonStrongNameSettingsFolderWithGreatestVersionNumber = individualDirectory.FullName

                End If

            Next

            'get the version number of the non strong name directory with the greatest version number

            Dim versionNumberOfNonStrongNameSettingsFolderWithGreatestVersionNumber As String = nonStrongNameSettingsFolderWithGreatestVersionNumber.Remove(0, nonStrongNameSettingsFolderWithGreatestVersionNumber.LastIndexOf("\") + 1)

            Dim newStrongNameSettingsFolderWithGreatestVersionNumber = parentOfStrongNameSettingsFolderForThisVersion & "\" & versionNumberOfNonStrongNameSettingsFolderWithGreatestVersionNumber

            If Directory.Exists(newStrongNameSettingsFolderWithGreatestVersionNumber) Then

                ' this should not happen

            Else

                '  copy over the old non strong name settings information to the new strong name folder so that it can be upgraded

                Directory.CreateDirectory(newStrongNameSettingsFolderWithGreatestVersionNumber)

                CopyDirectory(nonStrongNameSettingsFolderWithGreatestVersionNumber, newStrongNameSettingsFolderWithGreatestVersionNumber, True)


            End If

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub

    Private Sub CopyDirectory(sourceDir As String, destinationDir As String, recursive As Boolean)

        'the code for this subroutine was extracted from:
        'https://docs.microsoft.com/en-us/dotnet/standard/io/how-to-copy-directories

        ' Get information about the source directory
        Dim dir As New DirectoryInfo(sourceDir)

        ' Check if the source directory exists
        If Not dir.Exists Then
            Throw New DirectoryNotFoundException($"Source directory Not found:    {dir.FullName}")
        End If

        ' Cache directories before we start copying
        Dim dirs As DirectoryInfo() = dir.GetDirectories()

        ' Create the destination directory
        Directory.CreateDirectory(destinationDir)

        ' Get the files in the source directory and copy to the destination directory
        For Each file As FileInfo In dir.GetFiles()
            Dim targetFilePath As String = Path.Combine(destinationDir, file.Name)
            file.CopyTo(targetFilePath)
        Next

        ' If recursive and copying subdirectories, recursively call this method
        If recursive Then
            For Each subDir As DirectoryInfo In dirs
                Dim newDestinationDir As String = Path.Combine(destinationDir, subDir.Name)
                CopyDirectory(subDir.FullName, newDestinationDir, True)
            Next
        End If

    End Sub

#End Region

    Private MouseButtonIsUp As Boolean = True
    Private MouseDownMousePosition As New System.Drawing.Point

    Private Sub ListView1_PreviewMouseUp(sender As Object, e As MouseButtonEventArgs) Handles ListView1.PreviewMouseUp

        MouseButtonIsUp = True

    End Sub

    'v1.6
    Private Sub ListView1_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles ListView1.PreviewMouseDown

        Try

            If e.LeftButton = MouseButtonState.Pressed Then

                Dim pt As Point = e.GetPosition(Me)
                Dim ListViewComponent As Object = System.Windows.Media.VisualTreeHelper.HitTest(Me, pt)
                If ListViewComponent.visualhit.ToString = "System.Windows.Controls.Border" Then
                    ' mouse down event was on header row - ignore it
                Else
                    MouseButtonIsUp = False
                    MouseDownMousePosition = System.Windows.Forms.Cursor.Position
                    ListView1.Items.Item(0).Sortorder = 0
                End If
                ListViewComponent = Nothing
                pt = Nothing

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub ListView1_PreviewMouseMove(sender As Object, e As Input.MouseEventArgs) Handles ListView1.PreviewMouseMove

        'this code provides for a drag and drop of a .p2r out of the listview

        If MouseButtonIsUp Then Exit Sub

        ' Start Drag and Drop processing when the mouse is down and has moved at least five pixels in any direction
        If (Math.Abs(MouseDownMousePosition.X - System.Windows.Forms.Cursor.Position.X) > 5) OrElse (Math.Abs(MouseDownMousePosition.Y - System.Windows.Forms.Cursor.Position.Y) > 5) Then
            MouseButtonIsUp = True 'once we've got this far, set the MouseButtonIsUp flag to true so this code will not be processed again until the next time the mouse is down
        Else
            Exit Sub
        End If

        If (gCurrentlySelectedRow.Description = gMasterSwitch) OrElse (gCurrentlySelectedRow.Description = "") Then Exit Sub


        ' create the file to be dragged And dropped
        ' once staged For dropping, overwrite it For security purposes, And delete it

        Try

            Dim CurrentCard As CardStructure = InitalizeCard()

            With gCurrentlySelectedRow

                CurrentCard.Description = .Description
                CurrentCard.ListenFor = .ListenFor
                CurrentCard.Open = .Open
                CurrentCard.StartDirectory = .StartIn
                CurrentCard.Parameters = .Parameters
                CurrentCard.StartWithAdminPrivileges = .Admin
                CurrentCard.StartingWindowState = .StartingWindowState
                CurrentCard.KeysToSend = .KeysToSend

            End With

            Dim PathAndFileNameOfSavedCard As String = SaveCard(CurrentCard)

            'v4.3 saved filename will have underscores in place of spaces in the filename

            Dim PathNameOnly As String = Path.GetDirectoryName(PathAndFileNameOfSavedCard)
            Dim FileNameOnly As String = Path.GetFileName(PathAndFileNameOfSavedCard)

            PathAndFileNameOfSavedCard = PathNameOnly & "\" & FileNameOnly.Replace(" ", "_")

            Dim paths() As String = {PathAndFileNameOfSavedCard}

            DragDrop.DoDragDrop(Me, New System.Windows.DataObject(System.Windows.DataFormats.FileDrop, paths), System.Windows.DragDropEffects.Copy)

            If File.Exists(PathAndFileNameOfSavedCard) Then
                My.Computer.FileSystem.WriteAllText(PathAndFileNameOfSavedCard, StrDup(4096, "*"), False) ' overwrite file
                System.IO.File.Delete(paths(0)) ' delete it
            End If

            CurrentCard = Nothing

        Catch ex As Exception
        End Try

    End Sub

#End Region


    Private SharedObject = New Object

    Private Sub ClearFilters_Click(sender As Object, e As RoutedEventArgs) Handles ClearFilters.Click

        ClearAllFilters()
        SetLookOfMenus()

    End Sub

    Private Sub ClearAllFilters()

        tbFilterDescription.Text = String.Empty
        tbFilterListenFor.Text = String.Empty
        tbFilterOpen.Text = String.Empty
        tbFilterStartIn.Text = String.Empty
        tbFilterParameters.Text = String.Empty
        tbFilterAdmin.Text = String.Empty
        tbFilterStartingWindowState.Text = String.Empty
        tbFilterKeysToSend.Text = String.Empty

        FilterIsActive = False

    End Sub

    Private Sub FilterChanged(sender As Object, e As TextChangedEventArgs) Handles tbFilterDescription.TextChanged, tbFilterListenFor.TextChanged, tbFilterOpen.TextChanged, tbFilterStartIn.TextChanged, tbFilterParameters.TextChanged,
                                                                                   tbFilterAdmin.TextChanged, tbFilterStartingWindowState.TextChanged, tbFilterKeysToSend.TextChanged

        Static Dim NextTimeToRefresh As DateTime
        Static Dim CurrentFilter, LastFilter As String

        If tbFilterAdmin.Text.Length > 0 Then

            Dim ws = tbFilterAdmin.Text.Trim.ToLower
            Dim valid As Boolean = (("yes".StartsWith(ws)) OrElse ("no".StartsWith(ws)))

            If valid Then
            Else
                Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "Valid values for the Admin filter are 'Yes' and 'No'", gThisProgramName, MessageBoxButton.OK, MessageBoxImage.Asterisk, System.Windows.MessageBoxOptions.None)
                tbFilterAdmin.Text = tbFilterAdmin.Text.Remove(tbFilterAdmin.Text.Length - 1)
                tbFilterAdmin.Select(tbFilterAdmin.Text.Length, 0)
                Exit Sub
            End If

        End If

        If tbFilterStartingWindowState.Text.Trim.Length > 0 Then

            Dim ws = tbFilterStartingWindowState.Text.Trim.ToLower
            Dim valid As Boolean = (("minimized".StartsWith(ws)) OrElse ("normal".StartsWith(ws)) OrElse ("maximized".StartsWith(ws)) OrElse ("hidden".StartsWith(ws)))

            If valid Then
            Else
                Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "Valid values for the Window state are 'Minimized', 'Normal', 'Maximized' and 'Hidden'", gThisProgramName, MessageBoxButton.OK, MessageBoxImage.Asterisk, System.Windows.MessageBoxOptions.None)
                tbFilterStartingWindowState.Text = tbFilterStartingWindowState.Text.Remove(tbFilterStartingWindowState.Text.Length - 1)
                tbFilterStartingWindowState.Select(tbFilterStartingWindowState.Text.Length, 0)
                Exit Sub
            End If

        End If

        CurrentFilter = tbFilterDescription.Text.Trim & tbFilterListenFor.Text.Trim & tbFilterOpen.Text.Trim & tbFilterStartIn.Text.Trim & tbFilterParameters.Text.Trim & tbFilterAdmin.Text.Trim &
                        tbFilterStartingWindowState.Text.Trim & tbFilterKeysToSend.Text.Trim
        FilterIsActive = (CurrentFilter.Trim > String.Empty)

        'this logic saves refreshes when the user types quickly

        NextTimeToRefresh = Now.AddMilliseconds(100)

        While Now < NextTimeToRefresh
            Thread.Sleep(25)
            DoEvents()
        End While

        If CurrentFilter = LastFilter Then
            'saved a refresh
            Exit Sub
        Else
            LastFilter = CurrentFilter
            RefreshListView(CurrentSortOrder)
            sender.focus
        End If

        SetLookOfMenus()

    End Sub

#Region "Dropbox"

    Private Structure Struct_DropboxID
        Dim Name As String
        Dim Email As String
    End Structure

    Dim gDroboxID As Struct_DropboxID

    Private WithEvents DropboxWatcherFolder As FileSystemWatcher

    Private Sub SetupForDropBoxProcessing()

        Static CurrentlyWatching As String = String.Empty

        If My.Settings.UseDropbox Then

            DeleteDropBoxCommandFile()

            If CurrentlyWatching = My.Settings.DropboxPath Then
                Exit Sub
            End If

            Try

                If (DropboxWatcherFolder IsNot Nothing) Then

                    RemoveHandler DropboxWatcherFolder.Created, AddressOf DropboxTrigger

                    DropboxWatcherFolder.Dispose()
                    DropboxWatcherFolder = Nothing

                    Log("Dropbox monitoring ended for " & CurrentlyWatching)
                    Log("")

                    CurrentlyWatching = String.Empty

                End If

            Catch ex As Exception
                CurrentlyWatching = String.Empty
            End Try

            Try

                If Directory.Exists(My.Settings.DropboxPath.Trim) Then ' v 4.9

                    DropboxWatcherFolder = New System.IO.FileSystemWatcher()

                    'this is the path we want to monitor
                    DropboxWatcherFolder.Path = Settings.DropboxPath.Trim

                    'Add a list of Filter we want to specify
                    'make sure you use OR for each Filter as we need to
                    'all of those 

                    DropboxWatcherFolder.NotifyFilter = IO.NotifyFilters.DirectoryName
                    DropboxWatcherFolder.NotifyFilter = DropboxWatcherFolder.NotifyFilter Or IO.NotifyFilters.FileName
                    ' DropboxWatherFolder.NotifyFilter = DropboxWatherFolder.NotifyFilter Or IO.NotifyFilters.Attributes

                    ' add the handler to each event
                    AddHandler DropboxWatcherFolder.Created, AddressOf DropboxTrigger

                    'Set this property to true to start watching
                    DropboxWatcherFolder.EnableRaisingEvents = True

                    CurrentlyWatching = My.Settings.DropboxPath

                    Log("Dropbox monitoring begun for " & CurrentlyWatching)
                    Log("")

                Else

                    Log("Error: The Dropbox folder path, '" & Settings.DropboxPath.Trim & "', identified in the File - Options - Dropbox window doesn't exist")
                    Log("")

                End If

            Catch ex As Exception

                CurrentlyWatching = String.Empty
                Log(ex.ToString)

            End Try

        Else

            Try

                If (DropboxWatcherFolder IsNot Nothing) Then

                    RemoveHandler DropboxWatcherFolder.Created, AddressOf DropboxTrigger

                    DropboxWatcherFolder.Dispose()
                    DropboxWatcherFolder = Nothing

                    Log("Dropbox monitoring ended for " & CurrentlyWatching)
                    Log("")

                    CurrentlyWatching = String.Empty

                End If

            Catch ex As Exception
                CurrentlyWatching = String.Empty
            End Try


        End If

    End Sub

    Private Sub DropboxTrigger()

        If My.Settings.UseDropbox Then
        Else
            Exit Sub
        End If

        Try

            'Action Dropbox message

            Log("")
            Log("Incoming Dropbox trigger ...")

            Thread.Sleep(1000) ' give time for the file to be written

            Dim DropboxFullFilename As String = My.Settings.DropboxPath.Trim & My.Settings.DropboxFileName.Trim
            Dim DropboxFileNameWithoutExtention As String = Path.GetFileNameWithoutExtension(DropboxFullFilename)
            Dim DropboxExtention As String = Path.GetExtension(DropboxFullFilename)

            Dim files() As String = IO.Directory.GetFiles(My.Settings.DropboxPath.Trim)

            Dim ProcessThisFile As Boolean

            If files.Count > 0 Then

                For Each file As String In files

                    ProcessThisFile = False

                    If (file.ToUpper = DropboxFullFilename.ToUpper) Then

                        'match on Command.txt 
                        ProcessThisFile = True

                    ElseIf (Path.GetExtension(file).ToUpper = DropboxExtention.ToUpper) Then

                        If Path.GetFileNameWithoutExtension(file).ToUpper.StartsWith(DropboxFileNameWithoutExtention.ToUpper) Then

                            If Path.GetFileNameWithoutExtension(file).EndsWith(")") Then

                                'Match on Command (#).txt
                                ProcessThisFile = True

                            End If

                        End If

                    End If

                    If ProcessThisFile Then

                        Dim FileContents As String = ""
                        Try
                            FileContents = My.Computer.FileSystem.ReadAllText(file)
                        Catch ex As Exception
                        End Try

                        Log("Processing Dropbox file: " & file)
                        Log("File contents:" & vbCrLf & FileContents)

                        If FileContents.Trim.Length = 0 Then
                            Log("Incoming file was empty - no further processing action will be taken")
                            GoTo AllDoneForThisFile
                        End If

                        Dim Lines() As String
                        Lines = FileContents.Split(vbLf)

                        If ((Lines.Count < 2) OrElse (Lines.Count > 5)) OrElse (Lines(0).Trim.Length = 0) OrElse (Lines(1).Trim.Length = 0) Then

                            Log("Incoming file does not contain the expected information in the right format")
                            Log("The first line should identify the computer(s) for the command to run on")
                            Log("The second line should contain the command")
                            Log("For example:")
                            Log("   " & My.Settings.DropboxDeviceName)
                            Log("   open the calculator")
                            Dim dtdt As DateTime = Now
                            Dim dtstring = dtdt.ToString("MMMM dd, yyyy") & " at " & dtdt.ToString(" HH:mmtt")
                            Log(dtstring)
                            Log("No further action will be taken")
                            GoTo AllDoneForThisFile
                        End If

                        If (My.Settings.DropboxDeviceName.Length = 0) Then

                            Log("Dropbox device name not set in options, no further processing action will be taken")
                            GoTo AllDoneForThisFile

                        End If

                        Try

                            Dim DateInDropboxFile As DateTime = DateTime.Parse(Lines(2).Trim.Replace(" at ", " "))

                            If DateInDropboxFile < Now.AddSeconds(-180) Then
                                Log("The Dropbox file's ""Created At"" time is more than three minutes ago; the Dropbox file will be deleted and no further processing action will be taken")
                                Log("")
                                GoTo AllDoneForThisFile
                            End If

                        Catch ex As Exception

                        End Try

                        Dim DevicesIn() As String = Lines(0).Trim.ToUpper.Split(" ")
                        Dim DeviceInOptions As String = My.Settings.DropboxDeviceName.Trim.ToUpper
                        Dim MatchFound As Boolean = False

                        For Each DeviceName In DevicesIn
                            If DeviceName = DeviceInOptions Then
                                MatchFound = True
                                Exit For
                            End If
                        Next

                        If MatchFound Then

                            ActionIncomingMessage(MessageSource.Dropbox, Lines(1))
                            Thread.Sleep(6000) ' give time for other instances of Push2Run to react to the file

                        Else

                            Log("Device name in command file did not match Dropbox device name in options, no further processing action will be taken")

                        End If


AllDoneForThisFile:
                        Try
                            Log("Deleting : " & file)
                            My.Computer.FileSystem.DeleteFile(file)
                        Catch ex As Exception
                            Log("Delete failed")
                        End Try

                        'Log("")

                    Else

                        'Log("Ignoring file: " & file)

                    End If

                Next

            Else

                Log("Could not find any files to work with")
                'Log("")

            End If

            Log("")

        Catch ex As Exception

        End Try

    End Sub

    Private Sub DeleteDropBoxCommandFile()

        Try

            My.Computer.FileSystem.DeleteFile(My.Settings.DropboxPath.Trim & My.Settings.DropboxFileName.Trim)

        Catch ex As Exception
        End Try

    End Sub


#End Region


#Region "MQTT"

    'ref: https://github.com/dotnet/MQTTnet/wiki/Using-the-MQTT-client-in-VB.net

    Private ThreadQueue As Push2Run.Threading.SerialQueue = New Push2Run.Threading.SerialQueue

    Private Async Sub StartMQTTThread()

        Try

            If My.Settings.UseMQTT Then

                Await ThreadQueue.Enqueue(Async Function()

                                              Await MQTT_Connect(My.Settings.MQTTBroker, My.Settings.MQTTPort, My.Settings.MQTTUser, My.Settings.MQTTPassword)

                                              MQTTSetupComplete = True

                                              Return True

                                          End Function)


            End If


        Catch ex As Exception

            Log("Problem starting MQTT" & vbCrLf & ex.Message.ToString)

        End Try

        DoEvents()

    End Sub

    Enum SubscriptionAction
        LeaveSubscribed = 0
        Subscribe = 1
        Unsubscribe = 2
    End Enum

    Private Structure SubsciptionEntry
        Dim Action As SubscriptionAction
        Dim Topic As String
    End Structure

    Private ListOfSubscriptionActions As New List(Of SubsciptionEntry)
    Private ListOfCurrentlySubscribedTopics As New List(Of String)
    Private UpdateSubscriptionsComplete As Boolean
    Private Async Sub UpdateSubscriptions()

        UpdateSubscriptionsComplete = False

        Try

            ListOfSubscriptionActions.Clear()

            'Step 1: assume all old topics will be unsubscribed, this will be updated as needed in Step 3

            For Each OldTopic In ListOfCurrentlySubscribedTopics.ToList()

                Dim NewEntry As SubsciptionEntry

                With NewEntry
                    .Action = SubscriptionAction.Unsubscribe
                    .Topic = OldTopic
                End With

                ListOfSubscriptionActions.Add(NewEntry)

            Next

            If My.Settings.UseMQTT Then
            Else
                GoTo Step5
            End If

            'Step 2: build the list of new topics to subscribe to

            Dim Topics() As String = My.Settings.MQTTFilter.Split(" ")

            Dim LogEntry As String = "Current MQTT topics are:"

            For Each Topic In Topics.ToList()

                If Topic.Length > 0 Then

                    LogEntry &= " " & Topic.Trim

                End If

            Next

            Log(LogEntry)

            'Step 3: For each topic now listed in the options either mark it to be subscribed to or mark it to be left alone if it had already been subscribed to

            For Each Topic In Topics.ToList()

                If Topic.Length > 0 Then

                    Dim MatchFound As Boolean = False
                    For Each Entry In ListOfSubscriptionActions.ToList()

                        If Entry.Topic = Topic.Trim Then

                            MatchFound = True
                            Exit For

                        End If

                    Next

                    Dim OldEntry, NewEntry As SubsciptionEntry

                    NewEntry.Topic = Topic

                    If MatchFound Then

                        OldEntry.Topic = Topic
                        OldEntry.Action = SubscriptionAction.Unsubscribe
                        ListOfSubscriptionActions.Remove(OldEntry)

                        NewEntry.Action = SubscriptionAction.LeaveSubscribed

                    Else

                        NewEntry.Action = SubscriptionAction.Subscribe

                    End If

                    ListOfSubscriptionActions.Add(NewEntry)

                End If

            Next

            ' At this point the ListOfSubscriptionActions contains all topics to be subscribed, subscribed, or unsubscribed

            ' to make the session log easy to read actions will be processed based on how they appear in the new lists of topics, will all unsubscribing actions being done last


            'Step 4 process new subscriptions and leave alones 

            For Each Topic In Topics.ToList()

                For Each entry In ListOfSubscriptionActions.ToList()

                    If Topic = entry.Topic Then

                        If entry.Action = SubscriptionAction.LeaveSubscribed Then

                            Log(gThisProgramName & " is already subscribed to " & Topic)

                        Else

                            Await ThreadQueue.Enqueue(Async Function()

                                                          Await Subscribe(Topic, MQTTnet.Protocol.MqttQualityOfServiceLevel.ExactlyOnce)

                                                          Return True

                                                      End Function)

                        End If

                        Exit For

                    End If

                Next

            Next

Step5:

            ' Step 5 process unsubscribes

            For Each entry In ListOfSubscriptionActions.ToList()

                If entry.Action = SubscriptionAction.Unsubscribe Then

                    Await ThreadQueue.Enqueue(Async Function()

                                                  Await Unsubscribe(entry.Topic)

                                                  Return True

                                              End Function)

                End If

            Next


        Catch ex As Exception

            Log("Problem starting MQTT" & vbCrLf & ex.Message.ToString)

        End Try

        Log("")

        UpdateSubscriptionsComplete = True

        DoEvents()

    End Sub

    Private Async Sub UnsubscribeToAllCurrentSubscriptions()

        Try

            For Each Topic In ListOfCurrentlySubscribedTopics.ToList()

                Await ThreadQueue.Enqueue(Async Function()

                                              Await Unsubscribe(Topic)

                                              Return True

                                          End Function)


            Next

        Catch ex As Exception

        End Try

        Log("")

        DoEvents()

    End Sub

    '*********************

    Private Options As MqttClientOptionsBuilder
    Private Async Function MQTT_Connect(Server As String, Port As Integer, User As String, Password As String) As Task
        Try

            MQTTClient = MQTTFactory.CreateMqttClient()

            ' Set up event handlers using the new syntax
            AddHandler MQTTClient.DisconnectedAsync, AddressOf ConnectionClosed
            AddHandler MQTTClient.ApplicationMessageReceivedAsync, AddressOf MessageRecieved
            AddHandler MQTTClient.ConnectedAsync, AddressOf ConnectionOpened

            ' Create options using the new builder pattern
            Options = New MqttClientOptionsBuilder()

            Dim clientOptions = Options.WithClientId(gThisProgramName & "-" & Guid.NewGuid().ToString()) _
            .WithTcpServer(Server, Port) _
            .WithTimeout(TimeSpan.FromSeconds(10)) _
            .WithCredentials(User, EncryptionClass.Decrypt(Password)) _
            .Build()

            ' Connect using the new async method
            Await MQTTClient.ConnectAsync(clientOptions).ConfigureAwait(False)
        Catch ex As Exception
            Log("MQTT connection failed. " & ex.Message)
        End Try
    End Function

    Private Async Function MQTT_Disconnect() As Task
        Try
            Dim WaitUntil As DateTime = Now.AddSeconds(5)
            UnsubscribeToAllCurrentSubscriptions()

            While (ListOfCurrentlySubscribedTopics.Count > 0) AndAlso (Now < WaitUntil)
                System.Threading.Thread.Sleep(100)
                DoEvents()
            End While

            ' Use the new async disconnect method
            Await MQTTClient.DisconnectAsync().ConfigureAwait(False)
        Catch ex As Exception
            Log("MQTT disconnection issue: " & ex.Message)
        End Try
    End Function




    ' MQTT Publish is in modCommon

    Async Function Subscribe(Topic As String, ByVal QOS As MQTTnet.Protocol.MqttQualityOfServiceLevel) As Task

        Try

            Await MQTTClient.SubscribeAsync(Topic, QOS).ConfigureAwait(False)

            ListOfCurrentlySubscribedTopics.Add(Topic)

            Log(gThisProgramName & " subscribed to " & Topic)

            ' v4.1
            ' if we can subscribe then we are connected
            If gMQTTConnectionStatus = ConnectionStatus.Connected Then
            Else
                gMQTTConnectionStatus = ConnectionStatus.Connected
                gMQTTStatusText = "MQTT connected"
                ToastNotificationForNetorkEvents("MQTT", "Connection established", Now)
            End If

        Catch ex As Exception

            Log("Problem subscribing to " & Topic & vbCrLf & ex.Message.ToString)

        End Try

    End Function

    Async Function Unsubscribe(Topic As String) As Task

        Try

            Await MQTTClient.UnsubscribeAsync(Topic).ConfigureAwait(False)

            ListOfCurrentlySubscribedTopics.Remove(Topic)

            Log(gThisProgramName & " unsubscribed to " & Topic)

            ' v4.1
            ' if we can unsubscribe then we are connected; this is a safeguard   
            If gMQTTConnectionStatus = ConnectionStatus.Connected Then
            Else
                gMQTTConnectionStatus = ConnectionStatus.Connected
                gMQTTStatusText = "MQTT connected"
                ToastNotificationForNetorkEvents("MQTT", "Connection established", Now)
            End If


        Catch ex As Exception

            Log("Problem unsubscribing to " & Topic & vbCrLf & ex.Message.ToString)

        End Try

    End Function

    Private Async Function MessageRecieved(e As MqttApplicationMessageReceivedEventArgs) As Task

        Try

            Dim Topic As String = e.ApplicationMessage.Topic
            Dim Payload As String = System.Text.Encoding.UTF8.GetString(e.ApplicationMessage.PayloadSegment.ToArray()).Trim

            Log("MQTT incoming message:")
            Log("Topic:" & vbTab & Topic)
            Log("Payload:" & vbTab & Payload)
            Log("")

            ' v4.1
            ' if we are getting message than we are connected
            If gMQTTConnectionStatus = ConnectionStatus.Connected Then
            Else
                gMQTTConnectionStatus = ConnectionStatus.Connected
                gMQTTStatusText = "MQTT connected"
                ToastNotificationForNetorkEvents("MQTT", "Connection established", Now)
            End If

            If My.Settings.MQTTListenForPayloadOnly Then
            Else
                Payload = Topic & "/" & Payload
            End If

            ActionIncomingMessage(MessageSource.MQTT, Payload) 'v4.8

        Catch ex As Exception

            Log("MQTT message received issue:" & vbCrLf & ex.Message.ToString)
            Log("")

        End Try

        Return

    End Function

    Private Async Function ConnectionOpened(e As MqttClientConnectedEventArgs) As Task

        Log("MQTT connected")
        Log("")

        gMQTTConnectionStatus = ConnectionStatus.Connected
        gMQTTStatusText = "MQTT connected"
        ToastNotificationForNetorkEvents("MQTT", "Connection established", Now)

        Application.Current.Dispatcher.Invoke(New UpdateStatusBarNowCallback(AddressOf UpdateStatusBar))

        DoEvents()

    End Function

    Private Async Function ConnectionClosed(e As MqttClientDisconnectedEventArgs) As Task

        If gMQTTConnectionStatus = ConnectionStatus.DisconnectedorClosed Then

        Else

            ListOfCurrentlySubscribedTopics.Clear()

            Log("MQTT disconnected")
            Log("")

            gMQTTConnectionStatus = ConnectionStatus.DisconnectedorClosed
            gMQTTStatusText = "MQTT not connected"
            ToastNotificationForNetorkEvents("MQTT", "Connection not established", Now)

            Application.Current.Dispatcher.Invoke(New UpdateStatusBarNowCallback(AddressOf UpdateStatusBar))

            DoEvents()

        End If

    End Function

#End Region

End Class


<System.Reflection.ObfuscationAttribute(Feature:="renaming")> <XmlRootAttribute("MyTable1Class", [Namespace]:="", IsNullable:=False)>
Public Class MyTable1Class

    Public Property ID As Integer
    Public Property SortOrder As Integer
    Public Property DesiredStatus As Integer
    Public Property WorkingStatus As Integer
    Public Property Description As String
    Public Property ListenFor As String
    Public Property Open As String
    Public Property Parameters As String
    Public Property StartIn As String
    Public Property Admin As Boolean
    Public Property StartingWindowState As Integer
    Public Property KeysToSend As String

    Public Sub New(Optional ByVal ID As Integer = 0,
               Optional ByVal SortOrder As Integer = 0,
               Optional ByVal DesiredStatus As Integer = 0,
               Optional ByVal WorkingStatus As Integer = 0,
               Optional ByVal Description As String = "",
               Optional ByVal ListenFor As String = "",
               Optional ByVal Open As String = "",
               Optional ByVal Parameters As String = "",
               Optional ByVal StartIn As String = "",
               Optional ByVal Admin As Boolean = False,
               Optional ByVal StartingWindowState As Integer = 0,
               Optional ByVal KeysToSend As String = "")

        'Set fields
        Me.ID = ID
        Me.SortOrder = SortOrder
        Me.DesiredStatus = DesiredStatus
        Me.WorkingStatus = WorkingStatus
        Me.Description = Description
        Me.ListenFor = ListenFor
        Me.Open = Open
        Me.Parameters = Parameters
        Me.StartIn = StartIn
        Me.Admin = Admin
        Me.StartingWindowState = StartingWindowState
        Me.KeysToSend = KeysToSend

    End Sub

    Public Sub Initalize()
        'Set fields
        Me.ID = 0
        Me.SortOrder = 0
        Me.DesiredStatus = 0
        Me.WorkingStatus = 0
        Me.Description = String.Empty
        Me.ListenFor = String.Empty
        Me.Open = String.Empty
        Me.Parameters = String.Empty
        Me.StartIn = String.Empty
        Me.Admin = False
        Me.StartingWindowState = 0
        Me.KeysToSend = String.Empty
    End Sub

End Class

' The following Class is the same as the above class with two additional properties - DisplayableAdminText and DisplayableStartingWindowStateText

<System.Reflection.ObfuscationAttribute(Feature:="renaming")> <XmlRootAttribute("MyTable1ClassForTheListView", [Namespace]:="", IsNullable:=False)>
Public Class MyTable1ClassForTheListView

    Public Property ID As Integer
    Public Property SortOrder As Integer
    Public Property DesiredStatus As Integer
    Public Property WorkingStatus As Integer
    Public Property Description As String
    Public Property ListenFor As String
    Public Property Open As String
    Public Property Parameters As String
    Public Property StartIn As String
    Public Property Admin As Boolean
    Public Property StartingWindowState As Integer
    Public Property KeysToSend As String
    Public Property DisplayableAdminText As String
    Public Property DisplayableStartingWindowStateText As String

    Public Sub New(Optional ByVal ID As Integer = 0,
               Optional ByVal SortOrder As Integer = 0,
               Optional ByVal DesiredStatus As Integer = 0,
               Optional ByVal WorkingStatus As Integer = 0,
               Optional ByVal Description As String = "",
               Optional ByVal ListenFor As String = "",
               Optional ByVal Open As String = "",
               Optional ByVal Parameters As String = "",
               Optional ByVal StartIn As String = "",
               Optional ByVal Admin As Boolean = False,
               Optional ByVal StartingWindowState As Integer = 0,
               Optional ByVal KeysToSend As String = "",
               Optional ByVal DisplayableAdminText As String = "",
               Optional ByVal DisplayableStartingWindowStateText As String = "")

        'Set fields
        Me.ID = ID
        Me.SortOrder = SortOrder
        Me.DesiredStatus = DesiredStatus
        Me.WorkingStatus = WorkingStatus
        Me.Description = Description
        Me.ListenFor = ListenFor
        Me.Open = Open
        Me.Parameters = Parameters
        Me.StartIn = StartIn
        Me.Admin = Admin
        Me.StartingWindowState = StartingWindowState
        Me.KeysToSend = KeysToSend
        Me.DisplayableAdminText = DisplayableAdminText
        Me.DisplayableStartingWindowStateText = DisplayableStartingWindowStateText

    End Sub

    Public Sub Initalize()
        'Set fields
        Me.ID = 0
        Me.SortOrder = 0
        Me.DesiredStatus = 0
        Me.WorkingStatus = 0
        Me.Description = String.Empty
        Me.ListenFor = String.Empty
        Me.Open = String.Empty
        Me.Parameters = String.Empty
        Me.StartIn = String.Empty
        Me.Admin = False
        Me.StartingWindowState = 0
        Me.KeysToSend = String.Empty
        Me.DisplayableAdminText = String.Empty
        Me.DisplayableStartingWindowStateText = String.Empty
    End Sub

End Class
