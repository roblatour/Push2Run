'Copyright Rob Latour 2025

Imports System.Collections.ObjectModel
Imports System.Data
Imports System.Data.SQLite
Imports System.IO
Imports System.Net.Http
Imports System.Net.NetworkInformation
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.WindowsRuntime
Imports System.Security.Principal
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Windows.Automation
Imports System.Windows.Threading
Imports CommunityToolkit.WinUI.Notifications
Imports MQTTnet
Imports MQTTnet.Client
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Module modCommon

    '****************************************************************************************************************************

    '  **** very important to get the stuff below correct for a release  ****

    ' ensure the version in the visual studio project - properties - application settings is correct

    '                                                                 V
    Friend Const gSpecialProcessingForTesting As Boolean = False      ' False for production 
    Friend Const gBetaVersionInUse As Boolean = False                 ' False for production
    Friend Const PushbulletEnbledInTesting As Boolean = True          ' True for production  
    Friend Const PushoverEnabledInTesting As Boolean = True           ' True for production 
    '                                                                 ^

    ' Also update the file License-en.rtf with the current version number when building for release

    '  **** very important to get the stuff above correct for a release  ****

    '****************************************************************************************************************************

    Friend Const gThisProgramName As String = "Push2Run"

    Friend Const gGithubVersionControlPage As String = "https://raw.githubusercontent.com/roblatour/Push2Run/refs/heads/main/versionControl/"
    Friend gWebPageVersionCheck As String = gGithubVersionControlPage & "Push2RunCurrentVersion.txt"
    Friend gWebPageChangeLog As String = gGithubVersionControlPage & "Push2RunChangelog.rtf"
    Friend gAutomaticUpdateWebFileName As String = gGithubVersionControlPage & "Push2RunSetup.exe"

    Friend Const gLocalFiles As String = "file://E:\Documents\VBNet\Push2Run\Rackspace\"
    Friend Const gWebPageVersionCheckWhenTesting As String = gLocalFiles & "Push2RunCurrentVersion.txt"
    Friend Const gWebPageChangeLogWhenTesting As String = gLocalFiles & "Push2RunChangelog.rtf"
    Friend Const gAutomaticUpdateWebFileNameWhenTesting As String = gLocalFiles & "Push2RunSetup.exe"

    Friend Const gWebPageHomePage As String = "https://github.com/roblatour/Push2Run/"
    Friend Const gWebPageLicense As String = "https://github.com/roblatour/Push2Run/blob/main/LICENSE"

    Friend gWebPageHelp As String = gWebPageHomePage & "blob/main/help/help_vX.X.X.X.md" ' updated in code (WindowsBoss.loaded)
    Friend gWebPageHelpChangeWindow As String = String.Empty ' updated in code (WindowsBoss.loaded)
    Friend gWebPageHelpOptionsWindow As String = String.Empty ' updated in code (WindowsBoss.loaded)

    Friend Const gWebPageSetup As String = gWebPageHomePage & "blob/main/help/setup.md"

    Friend Const gWebPageDonate As String = "https://buymeacoffee.com/roblatour"
    Friend Const gWebPageDownload As String = gWebPageHomePage & "#download"
    Friend Const gProduceDetailDumpWhenNoMatchIsFound As Boolean = False

    Friend gAutomaticUpdateLocalDownloadedFileName As String = Path.GetTempPath & "Push2RunSetup.exe"

    Friend Const gAvailable As String = "available"
    Friend Const gNotAvailable As String = "not available"
    Friend Const g2FARequired As String = "2FA required"
    Friend Const gDefaultDropboxPath As String = "\Dropbox\Apps\Push2Run\"
    Friend Const gDefaultDropboxFilename As String = "Command.txt"

    Friend Const gPush2RunTriggerFileName As String = "Push2Run_admin_startup.txt"

    Friend Const gDefaultTimerIntervalInMilliseconds As Int32 = 1000 ' worker interval, every second, used to reload database, check if main window needs to be re-opened

    Friend Const gAboutHelpWindowTitle As String = "Push2Run - About/Help"

    'v4.6
    'Friend gSQLiteFullDatabaseName As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData) & "\Push2Run\Push2Run.db3"
    'Friend gSQLiteConnectionString As String = "Data Source=" & gSQLiteFullDatabaseName & ";"
    Friend gSQLiteFullDatabaseName As String
    Friend gSQLiteConnectionString As String

    Friend gSessionLogFile As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData) & "\Push2Run\SessionLog.txt"

    Friend Const gMasterSwitch = "Master Switch"
    Friend Const gWordToUseToDenoteAdminInListView = "Yes"

    Friend Const gTwelveHoursInMilliSeconds As Integer = 12 * 60 * 60 * 1000
    Friend Const gTwentyFourHoursInMilliSeconds As Integer = 24 * 60 * 60 * 1000

    Friend gCurrentOwner As Object
    Friend gNetworkMonitoringOn As Boolean = False

    Friend gSendingKeyesIsRequired As Boolean

    Friend gInitialStartupUnderway As Boolean = True
    Friend gBossLoadUnderway As Boolean = True
    Friend gShutdownUnderway As Boolean = False

    Friend gMenuSort As Boolean = True
    ' Friend gDatabaseWasCreated = False
    Friend Structure FileVersion
        Dim Major As Integer
        Dim Minor As Integer
        Dim Build As Integer
        Dim Revision As Integer
    End Structure

    Friend gCurrentlyRunningVersion As FileVersion
    Friend gCurrentVersionAccordingToWebsite As FileVersion

    Friend gVersionInUse As String
    Friend gWindowUpgradePromptIsOpen As Boolean = False

    Friend Enum RunningVersionRelativeToCurrentWebSiteVersion
        Unknown = 0
        RunningVersionIsNewerThanCurrentVersion = 1
        RunningVersionIsTheSameAsTheCurrentVersion = 2
        RunningVersionIsOlderThanTheCurrentVersion = 3
    End Enum

    Friend gRunningVersionRelativeToCurrentVersion As RunningVersionRelativeToCurrentWebSiteVersion

    Friend gThisProgramsPathAndFileName As String = String.Empty

    Friend Enum gOpenOptions
        AlwaysOnTop = 0
        Pushbullet = 1
        Pushover = 2
        Dropbox = 3
    End Enum

    Friend gOpenOptionsWindowAt As gOpenOptions = gOpenOptions.AlwaysOnTop

    Friend gMyUniqueID As String = String.Empty

    Friend LastSizeOfAddChangeWindow As Size = New Size(0, 0)
    Friend LastLocationOfAddChangeWindow As Point = New Point(-1, -1)

    Friend LastLocationofSessionLog As Point = New Size(0, 0)
    Friend LastSizeofSessionLog As Size = New Point(-1, -1)

    Friend LastTimeDataWasReceivedFromPushbullet As Date = Now.AddDays(-1)
    Friend LastTimeDataWasReceivedFromPushover As Date = Now.AddDays(-1)

    Friend gCriticalPushOverErrorReported As Boolean = False

    Friend gIgnoreAction As Boolean = False

    Friend Base As Object = Application.Current
    Friend Enum WorkerStatus
        Running = 1
        Stopped = 2
        Paused = 3
        Requested = 4
        NotYetKnown = 5
    End Enum

    Enum ActionStatus
        Succeeded = 0
        PartiallySucceeded = 1
        MasterSwitchWasOff = 2
        NotProcessedAsNoMatchingPhrasesFound = 3
        NotProcessedWhileAtLeastOneMatchingPhraseWasFoundNoneWereEnabled = 4
        NotProcecessNoProgramToRun = 5
        NotProcecessAsAUACPromptWouldBeRequired = 6
        Failed = 7
        Unknown = 8
        LeaveBlank = 9
        NotYetSet = 10
    End Enum

    Friend Structure gCommandRecord
        Friend ID As Integer
        ' Description is not needed in the command table because there are no processing decisions are made based on it
        Friend ListenFor As String
        Friend Open As String
        Friend Parameters As String
        Friend StartIn As String
        Friend Admin As Boolean
        Friend StartingWindowState As Integer
        Friend KeysToSend As String
        Friend DesiredStatus As StatusValues ' added in v3.4
    End Structure
    Friend Structure gAddEditRecord
        Friend DesiredStatus As StatusValues ' this is not used on the add entry window, but is used in the import function
        Friend Description As String
        Friend ListenFor As String
        Friend Open As String
        Friend Parameters As String
        Friend StartIn As String
        Friend Admin As Boolean
        Friend StartingWindowState As Integer
        Friend KeysToSend As String
    End Structure

    Friend Structure gImportExportRecord
        Friend Description As String
        Friend ListenFor As String
        Friend Open As String
        Friend Parameters As String
        Friend StartIn As String
        Friend Admin As Boolean
        Friend StartingWindowState As Integer
        Friend KeysToSend As String
    End Structure

    Friend NewAddEditRecord As gAddEditRecord
    Friend Function InitalizeAddEditRecord() As gAddEditRecord

        Dim ReturnValue As New gAddEditRecord

        With ReturnValue
            .Description = String.Empty
            .ListenFor = String.Empty
            .Open = String.Empty
            .Parameters = String.Empty
            .StartIn = String.Empty
            .Admin = vbFalse
            .StartingWindowState = 0
            .KeysToSend = String.Empty
        End With

        Return ReturnValue

    End Function

    Friend gPushoverAuthorizationReturnCode As String = String.Empty ' valid values: 'available', 'not available', '2FA required', string.empty
    Friend PushoverDeviceNameAndId As String = String.Empty ' valid values: 'available', 'not available', string.empty
    Friend PushoverMessageIdToDelete As String = String.Empty ' valid values: a number as string, string.empty
    Friend FindHighestPushoverMessageId_Processing As String = String.Empty ' valid values: 'done', string.empty
    Friend DeletePushoverMessageID_Processing As String = String.Empty ' valid values: 'done', string.empty
    Friend Enum StatusValues
        SwitchOn = 1
        SwitchOff = 2
        NoSwitch = 3
    End Enum
    Friend SwitchValues As StatusValues

    Friend Const MaxNumberOfEntries As Int32 = 1000
    Friend CommandTable(MaxNumberOfEntries) As gCommandRecord
    Friend CommandLineMaxIndex As Int32 = MaxNumberOfEntries

    Friend AllEntriesTable(MaxNumberOfEntries) As gImportExportRecord
    Friend AllEntiresMaxIndex As Int32 = MaxNumberOfEntries

    Friend Enum UpdateCheckFrequency
        Daily = 1
        Weekly = 2
        EveryTwoWeeks = 3
        Monthly = 4
    End Enum
    Friend Enum MonitorStatus
        Running = 1
        Stopped = 2
    End Enum
    Public Structure gRowRecord

        'list view rows
        Public ID As Long
        Public SortOrder As Integer
        Public DesiredStatus As Integer
        Public WorkingStatus As Integer
        Public Description As String
        Public ListenFor As String
        Public Open As String
        Public Parameters As String
        Public StartIn As String
        Public Admin As Boolean
        Public StartingWindowState As Integer
        Public KeysToSend As String

    End Structure

    Friend gCurrentlySelectedRow As gRowRecord

    Friend gMasterStatus As MonitorStatus = MonitorStatus.Running

    Friend gLoadingAddChange As String = String.Empty
    Friend gReturnFromAddChange As String = String.Empty
    Friend gReturnFromAddChangeDataChanged As Boolean = False
    Friend gDropInProgress As Boolean = False

    Friend gPasswordWasCorrectlyEnteredInPasswordWindow As Boolean = False
    Friend g2FAWasCorrectlyEnteredIn2FAWindow As Boolean = False
    Friend gPasswordWasCorrectlyEnteredInPasswordWindow_UserClickedCancel As Boolean = False

    Friend gEnterPasswordWindowTitle As String = String.Empty
    Friend gEnteredPassword As String = String.Empty
    Friend gWaitingForAResponse As Boolean = False
    Friend gEnteredPlainText2FA As String = String.Empty

    Friend DragAndDropUnderway As Boolean = False
    Friend IndexToBeSelectedOnReload As Integer = -1
    Friend ItemContentToBeScrolledToOnReload As String = String.Empty

    Friend gWindowUpgradePrompt As WindowUpgradePrompt
    Friend ghCurrentOwner As New Window

    Friend Const gMasterSwitchID As Integer = 1
    Friend Const gMasterSwitchSortOrder As Integer = 0 ' would have liked 0; but 1 is used for backwards compatibility with older versions of Push2Run
    Friend Const gGapBetweenSortIDsForDatabaseEntries As Integer = 2 ^ 13 ' this will allow for up to 10 inserts between records before a table rebuild is required

    Friend gHandleOfActiveWindow As IntPtr = 0

    Friend Enum P2RFileType
        OriginalStyeContainsOneCard = 1
        ExportStyleMayContainMultipleCards = 2
    End Enum

    Friend Enum ConnectionStatus
        Unknown = 0
        Connected = 1
        DisconnectedorClosed = 2
    End Enum

    Friend gPreviousNetworkConnectionStatus As OperationalStatus = OperationalStatus.Unknown

    Friend gPushoverConnectionStatus As ConnectionStatus = ConnectionStatus.Unknown
    Friend gPushbulletConnectionStatus As ConnectionStatus = ConnectionStatus.Unknown
    Friend gMQTTConnectionStatus As ConnectionStatus = ConnectionStatus.Unknown
    Friend gMQTTStatusText As String = String.Empty

    Friend MQTTSetupComplete As Boolean = False

    Friend gHandle As IntPtr


    'v4.6

    Private lockingobj As New Object

    Function IsFileAvailable(ByVal filename As String, Optional ByVal WaitTimeInMilliSeconds As Integer = 500, Optional ByVal RestTimeInMilliSeconds As Integer = 100) As Boolean

        SyncLock lockingobj

            Dim ReturnCode As Boolean = False

            Dim MaxWaitTime As DateTime = Now.AddMilliseconds(WaitTimeInMilliSeconds)

            Try

                'first check that the file exists; if it doesn't exist then there is no point checking to see if it is available

                If File.Exists(filename) Then
                    ' keep going
                Else
                    Exit Try
                End If

                ' Second check to see if the file is available; if it is not immediately available wait for up to max wait time for it to become available

                Do

                    Try

                        Dim FileNum As Integer = FreeFile()
                        FileOpen(FileNum, filename, OpenMode.Binary, OpenAccess.Read, OpenShare.LockReadWrite)
                        FileClose(FileNum)
                        ReturnCode = True
                        Exit Do

                    Catch ex As Exception
                        ' Waiting for file to become available
                    End Try

                    DoEvents()
                    If RestTimeInMilliSeconds > 0 Then
                        Thread.Sleep(RestTimeInMilliSeconds)
                    End If

                Loop Until (Now >= MaxWaitTime)

            Catch ex As Exception
            End Try

            Return ReturnCode

        End SyncLock

    End Function



    Friend Function DetermineP2RFileType(ByVal filename As String) As P2RFileType

        Dim AllText As String = My.Computer.FileSystem.ReadAllText(filename)

        If AllText.StartsWith("[") Then
            Return P2RFileType.ExportStyleMayContainMultipleCards
        Else
            Return P2RFileType.OriginalStyeContainsOneCard
        End If

    End Function

    Friend Function IsPushOverFullyConfigured() As Boolean

        Return ArePushoverIDAndSecretAvailable() AndAlso ArePushoverDeviceIDAndSecretAvailable() AndAlso ArePushoverDeviceNameAndIDAvailable()

    End Function

    Friend Function FolderAndFileCombine(ByVal Folder As String, ByVal File As String) As String

        Folder = Folder.Trim.TrimEnd("/").TrimEnd("\")
        File = File.Trim.TrimStart("/").TrimStart("\")
        Return Folder & "/" & File

    End Function

    Friend Sub rtbReplace(ByRef RTB As RichTextBox, ByVal oldString As String, newString As String)

        Dim text As TextRange = New TextRange(RTB.Document.ContentStart, RTB.Document.ContentEnd)
        Dim current As TextPointer = text.Start.GetInsertionPosition(LogicalDirection.Forward)

        While current IsNot Nothing

            Dim textInRun As String = current.GetTextInRun(LogicalDirection.Forward)

            If Not String.IsNullOrWhiteSpace(textInRun) Then

                Dim index As Integer = textInRun.IndexOf(oldString)

                If index <> -1 Then
                    Dim selectionStart As TextPointer = current.GetPositionAtOffset(index, LogicalDirection.Forward)
                    Dim selectionEnd As TextPointer = selectionStart.GetPositionAtOffset(oldString.Length, LogicalDirection.Forward)
                    Dim selection As TextRange = New TextRange(selectionStart, selectionEnd)
                    selection.Text = newString
                    'selection.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Bold)
                    'RTB.Selection.[Select](selection.Start, selection.[End])
                    'RTB.Focus()
                End If
            End If

            current = current.GetNextContextPosition(LogicalDirection.Forward)

        End While

    End Sub



    Friend Sub LoadAllEntriesFromDatabase()

        ' loads all entries into a table used by the import function to ensure duplicate entries are not imported
        ' this routine is much like LoadFromDatabase, but has some differences

        'Make sure the command database file exists
        If File.Exists(gSQLiteFullDatabaseName) Then
        Else
            For z As Int32 = 1 To 3
                Beep()
                System.Threading.Thread.Sleep(500)
            Next
            Dim currentProcess As Process = Process.GetCurrentProcess()
            currentProcess.Kill()
        End If

        'Load the command table from the database

        On Error Resume Next

        Dim Description As String
        Dim WorkingID As Integer
        Dim DesiredStatus As Integer
        Dim ListenFor As String
        Dim Open As String
        Dim Parameters As String
        Dim StartIn As String
        Dim Admin As Boolean
        Dim StartingWindowState As Integer
        Dim KeysToSend As String

        Dim sSQL2 As String = "SELECT * FROM Table1 ORDER BY SortOrder ASC ;"
        Dim SQLiteConnect2 As New SQLiteConnection(gSQLiteConnectionString)
        Dim SQLiteCommand2 As SQLiteCommand = New SQLite.SQLiteCommand(sSQL2, SQLiteConnect2)

        SQLiteConnect2.Open()

        Dim SQLiteDataReader2 As SQLiteDataReader = SQLiteCommand2.ExecuteReader(CommandBehavior.CloseConnection)

        AllEntiresMaxIndex = 100
        ReDim AllEntriesTable(AllEntiresMaxIndex)

        Dim x As Int32 = 0

        While SQLiteDataReader2.Read()

            Description = String.Empty
            DesiredStatus = StatusValues.SwitchOff
            ListenFor = String.Empty
            Open = String.Empty
            Parameters = String.Empty
            StartIn = String.Empty
            Admin = False
            StartingWindowState = 0
            KeysToSend = String.Empty

            DesiredStatus = SQLiteDataReader2.GetInt32(DatabaseColumns.DesiredStatus)

            'make sure the table is big enough to store all entries
            If x = AllEntiresMaxIndex Then
                AllEntiresMaxIndex += 100
                ReDim Preserve AllEntriesTable(AllEntiresMaxIndex)
            End If

            If DesiredStatus = StatusValues.NoSwitch Then
                ' Don't load blank lines

            Else

                Description = SQLiteDataReader2.GetString(DatabaseColumns.Description)
                ListenFor = SQLiteDataReader2.GetString(DatabaseColumns.ListenFor)  ' Defer decrypting info until the code below
                Open = SQLiteDataReader2.GetString(DatabaseColumns.Open) ' Defer decrypting info until the code below
                Parameters = SQLiteDataReader2.GetString(DatabaseColumns.Parameters) ' Defer decrypting info until the code below
                StartIn = SQLiteDataReader2.GetString(DatabaseColumns.StartIn) ' Defer decrypting info until the code below
                Admin = SQLiteDataReader2.GetValue(DatabaseColumns.Admin)
                StartingWindowState = SQLiteDataReader2.GetInt32(DatabaseColumns.StartingWindowState)
                KeysToSend = SQLiteDataReader2.GetString(DatabaseColumns.KeysToSend) ' Defer decrypting info until the code below

                With AllEntriesTable(x)
                    .Description = EncryptionClass.Decrypt(Description)
                    .ListenFor = EncryptionClass.Decrypt(ListenFor)
                    .Open = EncryptionClass.Decrypt(Open)
                    .Parameters = EncryptionClass.Decrypt(Parameters)
                    .StartIn = EncryptionClass.Decrypt(StartIn)
                    .Admin = Admin
                    .StartingWindowState = StartingWindowState
                    .KeysToSend = EncryptionClass.Decrypt(KeysToSend)
                End With

                x += 1

            End If

        End While

        SQLiteDataReader2.Close()
        SQLiteConnect2.Close()

        SQLiteConnect2.Dispose()
        SQLiteCommand2.Dispose()
        SQLiteDataReader2 = Nothing


    End Sub


    Friend Sub LoadFromDatabase(Optional ByVal ID As Integer = -1)

        ' the following loads the command table from the SQL database
        ' of note the description is not needed in the command table because there are no processing decisions are made based on it

        'Dim ReloadEntireTable As Boolean = MenuSort.IsChecked OrElse (ID = -1)  'v2.5.3 using gMenuSort to avoid cross thread issue
        Dim ReloadEntireTable As Boolean = gMenuSort OrElse (ID = -1)

        'Make sure the command database file exists
        If File.Exists(gSQLiteFullDatabaseName) Then
        Else
            For z As Int32 = 1 To 3
                Beep()
                System.Threading.Thread.Sleep(500)
            Next
            Dim currentProcess As Process = Process.GetCurrentProcess()
            currentProcess.Kill()
        End If

        'Load the command table from the database

        On Error Resume Next

        Dim WorkingID As Integer
        Dim DesiredStatus As Integer
        Dim ListenFor As String
        Dim Open As String
        Dim Parameters As String
        Dim StartIn As String
        Dim Admin As Boolean
        Dim StartingWindowState As Integer
        Dim KeysToSend As String

        Dim sSQL2 As String = "SELECT * FROM Table1 ORDER BY SortOrder ASC ;"
        Dim SQLiteConnect2 As New SQLiteConnection(gSQLiteConnectionString)
        Dim SQLiteCommand2 As SQLiteCommand = New SQLite.SQLiteCommand(sSQL2, SQLiteConnect2)

        SQLiteConnect2.Open()

        Dim SQLiteDataReader2 As SQLiteDataReader = SQLiteCommand2.ExecuteReader(CommandBehavior.CloseConnection)

        If ReloadEntireTable Then
            CommandLineMaxIndex = 100
            ReDim CommandTable(CommandLineMaxIndex)
        End If

        Dim x As Int32 = 0

        While SQLiteDataReader2.Read()

            DesiredStatus = StatusValues.SwitchOff
            ListenFor = String.Empty
            Open = String.Empty
            Parameters = String.Empty
            StartIn = String.Empty
            Admin = False
            StartingWindowState = 0
            KeysToSend = String.Empty

            WorkingID = SQLiteDataReader2.GetInt32(DatabaseColumns.ID)
            DesiredStatus = SQLiteDataReader2.GetInt32(DatabaseColumns.DesiredStatus)
            ListenFor = SQLiteDataReader2.GetString(DatabaseColumns.ListenFor)  ' Defer decrypting info until the code below (2 separate places) when we know it will be needed
            Open = SQLiteDataReader2.GetString(DatabaseColumns.Open) ' Defer decrypting info until the code below (2 separate places) when we know it will be needed
            Parameters = SQLiteDataReader2.GetString(DatabaseColumns.Parameters) ' Defer decrypting info until the code below (2 separate places) when we know it will be needed
            StartIn = SQLiteDataReader2.GetString(DatabaseColumns.StartIn) ' Defer decrypting info until the code below (2 separate places) when we know it will be needed
            Admin = SQLiteDataReader2.GetValue(DatabaseColumns.Admin)
            StartingWindowState = SQLiteDataReader2.GetInt32(DatabaseColumns.StartingWindowState)
            KeysToSend = SQLiteDataReader2.GetString(DatabaseColumns.KeysToSend) ' Defer decrypting info until the code below (2 separate places) when we know it will be needed

            If ReloadEntireTable Then

                'only load table entries where monitoring is turned on and there is something to do 
                ' If (DesiredStatus = StatusValues.SwitchOn) AndAlso (Open.Trim.Length > 0) Then

                If ((DesiredStatus = StatusValues.SwitchOn) OrElse (DesiredStatus = StatusValues.SwitchOff)) AndAlso (Open.Trim.Length > 0) Then ' updated in v3.4 to include entries that are switched off

                    'make sure the table is big enough to store all entries
                    If x = CommandLineMaxIndex Then
                        CommandLineMaxIndex += 100
                        ReDim Preserve CommandTable(CommandLineMaxIndex)
                    End If

                    CommandTable(x).ID = WorkingID
                    CommandTable(x).ListenFor = EncryptionClass.Decrypt(ListenFor)
                    CommandTable(x).Open = EncryptionClass.Decrypt(Open)
                    CommandTable(x).Parameters = EncryptionClass.Decrypt(Parameters)
                    CommandTable(x).StartIn = EncryptionClass.Decrypt(StartIn)
                    CommandTable(x).Admin = Admin
                    CommandTable(x).StartingWindowState = StartingWindowState
                    CommandTable(x).KeysToSend = EncryptionClass.Decrypt(KeysToSend)
                    CommandTable(x).DesiredStatus = DesiredStatus

                    x += 1

                End If

            Else

                If WorkingID = ID Then

                    If DesiredStatus = StatusValues.SwitchOn Then

                        'Look for the entry in the control table
                        'if found overwrite current entry
                        'otherwise add a new entry to the bottom of the table

                        Dim NewTableIndex As Int32 = -1
                        For y As Int32 = 0 To CommandTable.Length - 1
                            If CommandTable(y).ID = ID Then
                                NewTableIndex = y
                                Exit For
                            End If
                        Next

                        If NewTableIndex = -1 Then
                            ReDim Preserve CommandTable(CommandTable.Length)
                            NewTableIndex = CommandTable.Length - 1
                        End If

                        CommandTable(NewTableIndex).ID = WorkingID
                        CommandTable(NewTableIndex).ListenFor = EncryptionClass.Decrypt(ListenFor)
                        CommandTable(NewTableIndex).Open = EncryptionClass.Decrypt(Open)
                        CommandTable(NewTableIndex).Parameters = EncryptionClass.Decrypt(Parameters)
                        CommandTable(NewTableIndex).StartIn = EncryptionClass.Decrypt(StartIn)
                        CommandTable(NewTableIndex).Admin = Admin
                        CommandTable(NewTableIndex).StartingWindowState = StartingWindowState
                        CommandTable(NewTableIndex).KeysToSend = EncryptionClass.Decrypt(KeysToSend)

                    Else

                        'remove from table
                        Dim NewTableIndex As Int32 = -1
                        For y As Int32 = 0 To CommandTable.Length - 1
                            If CommandTable(y).ID = ID Then
                                NewTableIndex = y
                                Exit For
                            End If
                        Next

                        If NewTableIndex = -1 Then Exit While 'entry didn't exist

                        For y As Int32 = NewTableIndex + 1 To CommandTable.Length - 1
                            CommandTable(y - 1) = CommandTable(y)
                        Next

                        ReDim Preserve CommandTable(CommandTable.Length - 2)

                    End If

                    Exit While

                End If

            End If

        End While

        SQLiteDataReader2.Close()
        SQLiteConnect2.Close()

        SQLiteConnect2.Dispose()
        SQLiteCommand2.Dispose()
        SQLiteDataReader2 = Nothing

        If ReloadEntireTable Then ReDim Preserve CommandTable(x - 1)

    End Sub
    Friend Function SearchForAFileInTheSystemPath(ByVal Filename As String) As Boolean

        Dim ReturnCode As Boolean = False

        Try

            Dim CombinedPath As String = String.Empty

            Dim AllPaths = Environment.GetEnvironmentVariable("PATH")
            For Each folder As String In AllPaths.Split(";")

                CombinedPath = Path.Combine(folder.Trim, Filename).ToString

                'Console.WriteLine(CombinedPath)

                If File.Exists(CombinedPath) Then
                    ReturnCode = True
                    Exit For
                End If

            Next

        Catch ex As Exception
        End Try

        Return ReturnCode

    End Function
    Friend Function ArePushoverIDAndSecretAvailable() As Boolean

        Dim ReturnValue As Boolean = True

        Try

            If (My.Settings.PushoverID = gNotAvailable) OrElse
               (My.Settings.PushoverSecret = gNotAvailable) OrElse
               (EncryptionClass.Decrypt(My.Settings.PushoverID) = String.Empty) OrElse
               (EncryptionClass.Decrypt(My.Settings.PushoverSecret) = String.Empty) Then
                ReturnValue = False
            End If

        Catch ex As Exception
            ReturnValue = False
        End Try

        Return ReturnValue

    End Function

    Friend Function ArePushoverDeviceIDAndSecretAvailable() As Boolean

        Dim ReturnValue As Boolean = True

        Try

            If (My.Settings.PushoverDeviceID = gNotAvailable) OrElse
               (My.Settings.PushoverSecret = gNotAvailable) OrElse
               (EncryptionClass.Decrypt(My.Settings.PushoverDeviceID) = String.Empty) OrElse
               (EncryptionClass.Decrypt(My.Settings.PushoverSecret) = String.Empty) Then
                ReturnValue = False
            End If

        Catch ex As Exception
            ReturnValue = False
        End Try

        Return ReturnValue

    End Function

    Friend Function ArePushoverDeviceNameAndIDAvailable() As Boolean

        Dim ReturnValue As Boolean = True

        Try

            If (My.Settings.PushoverDeviceName = gNotAvailable) OrElse
               (My.Settings.PushoverDeviceID = gNotAvailable) OrElse
               (My.Settings.PushoverDeviceName = String.Empty) OrElse
               (EncryptionClass.Decrypt(My.Settings.PushoverDeviceID) = String.Empty) Then
                ReturnValue = False
            End If

        Catch ex As Exception
            ReturnValue = False
        End Try

        Return ReturnValue

    End Function

    Private Sub AuthorizeViaPushover(ByVal email As String, ByVal password As String, ByVal twoFactorAuthenticationCode As String)

        Dim Parms As New PushoverIdAndSecretParmClass
        Parms.email = email
        Parms.password = password
        Parms.twoFactorAuthenticationCode = twoFactorAuthenticationCode

        gPushoverAuthorizationReturnCode = String.Empty

        Dim NewThread As Thread = New Thread(AddressOf SetPushoverIDandSecret_Async)
        NewThread.Start(Parms)

        Thread.Sleep(250)
        While gPushoverAuthorizationReturnCode = String.Empty
            Thread.Sleep(100)
        End While

        NewThread = Nothing

    End Sub

    'v4.8.3 added Pushover 2FA support
    Friend Sub SetPushoverIdAndSecret(ByVal email As String, ByVal password As String)

        Try

            Log("Pushover authorization required")
            Log("")

            Dim twoFactorAuthenticationCode As String = String.Empty

            AuthorizeViaPushover(email, password, twoFactorAuthenticationCode)

            If gPushoverAuthorizationReturnCode = g2FARequired Then

                Log("2FA code required")
                Log("")

                'have the user enter the 2FA now
                Dim WindowPromptFor2FAWindow As WindowPromptFor2FA = New WindowPromptFor2FA

                gCurrentOwner = WindowPromptFor2FAWindow
                WindowPromptFor2FAWindow.ShowDialog()
                gCurrentOwner = Application.Current.MainWindow

                If g2FAWasCorrectlyEnteredIn2FAWindow Then

                    AuthorizeViaPushover(email, password, gEnteredPlainText2FA)

                    If gPushoverAuthorizationReturnCode = gAvailable Then

                        Log("2FA confirmed")
                        Log("Pushover authorized")
                        Log("")

                    End If

                    gEnteredPlainText2FA = ""

                End If

                If gPushoverAuthorizationReturnCode = gNotAvailable Then
                    Log("2FA not confirmed - Pushover authorization failed")
                    Log("")
                End If

            ElseIf gPushoverAuthorizationReturnCode = gAvailable Then

                Log("Pushover authorized")
                Log("")

            ElseIf gPushoverAuthorizationReturnCode = gNotAvailable Then

                Log("Pushover authorization failed")
                Log("")

            End If

        Catch ex As Exception
            gPushoverAuthorizationReturnCode = gNotAvailable

        End Try

    End Sub

    Private Async Sub SetPushoverIDandSecret_Async(ByVal data As Object)

        gPushoverAuthorizationReturnCode = String.Empty

        Dim ReturnValue As String = String.Empty

        Try

            Dim Parms = CType(data, PushoverIdAndSecretParmClass)

            Using httpClient = New HttpClient()

                Using request = New HttpRequestMessage(New HttpMethod("POST"), "https://api.pushover.net/1/users/login.json")

                    Dim multipartContent = New MultipartFormDataContent()
                    multipartContent.Add(New StringContent(Parms.email), "email")
                    multipartContent.Add(New StringContent(Parms.password), "password")
                    If Parms.twoFactorAuthenticationCode <> String.Empty Then
                        multipartContent.Add(New StringContent(Parms.twoFactorAuthenticationCode), "twofa")
                    End If
                    request.Content = multipartContent


                    Dim response As HttpResponseMessage = Await httpClient.SendAsync(request)

                    Dim responseBody As String = Await response.Content.ReadAsStringAsync()

                    Dim Pushover_Status As String = GetFirstMatchingValueFromJSONResponseString("status", responseBody)

                    If Pushover_Status = "1" Then

                        Dim Pushover_Secret As String = GetFirstMatchingValueFromJSONResponseString("secret", responseBody)
                        Dim Pushover_Id As String = GetFirstMatchingValueFromJSONResponseString("id", responseBody)

                        If (My.Settings.PushoverID <> EncryptionClass.Encrypt(Pushover_Id)) OrElse
                           (My.Settings.PushoverSecret <> EncryptionClass.Encrypt(Pushover_Secret)) Then

                            My.Settings.PushoverID = EncryptionClass.Encrypt(Pushover_Id)
                            My.Settings.PushoverSecret = EncryptionClass.Encrypt(Pushover_Secret)
                            My.Settings.Save()

                        End If

                        ReturnValue = gAvailable


                    ElseIf Pushover_Status = "0" Then

                        Dim twoFACheck As String = GetFirstMatchingValueFromJSONResponseString("totp", responseBody)
                        If twoFACheck = "two factor auth required" Then
                            ReturnValue = g2FARequired
                        Else
                            ReturnValue = gNotAvailable
                        End If

                    Else
                        ReturnValue = gNotAvailable
                    End If

                End Using

            End Using

        Catch ex As Exception

            ReturnValue = gNotAvailable

        End Try

        gPushoverAuthorizationReturnCode = ReturnValue

    End Sub

    Friend Sub SetPushoverDeviceNameAndID(ByVal DeviceName As String)

        Dim Parms As New PushoverDeviceNameClass

        Parms.name = DeviceName
        Parms.force = ForceValues.NoForce

        PushoverDeviceNameAndId = String.Empty
        Dim NewThread As Thread = New Thread(AddressOf SetPushoverDeviceNameAndID_Async)
        NewThread.Start(Parms)

        Thread.Sleep(250)
        While PushoverDeviceNameAndId = String.Empty
            Thread.Sleep(100)
        End While

        If PushoverDeviceNameAndId = gNotAvailable Then
            Log("Could not setup Pushover device")
            Log("")
        End If

        NewThread = Nothing

    End Sub

    Private Async Sub SetPushoverDeviceNameAndID_Async(ByVal data As Object)

        PushoverDeviceNameAndId = String.Empty

        Dim ReturnValue As String = String.Empty

        Try

            Dim Parms = CType(data, PushoverDeviceNameClass)

            If Parms.force = ForceValues.ForceApplied Then
                ReturnValue = gNotAvailable
                Exit Sub
            End If

            Using httpClient = New HttpClient()

                Using request = New HttpRequestMessage(New HttpMethod("POST"), "https://api.pushover.net/1/devices.json")

                    Dim multipartContent = New MultipartFormDataContent()

                    multipartContent.Add(New StringContent(EncryptionClass.Decrypt(My.Settings.PushoverSecret)), "secret")
                    multipartContent.Add(New StringContent(Parms.name), "name")
                    multipartContent.Add(New StringContent("O"), "os")

                    If Parms.force = ForceValues.Force Then
                        multipartContent.Add(New StringContent("1"), "force")
                        Parms.force = ForceValues.ForceApplied
                    End If

                    request.Content = multipartContent

                    Dim response As HttpResponseMessage = Await httpClient.SendAsync(request)
                    '  response.EnsureSuccessStatusCode()

                    Dim responseBody As String = Await response.Content.ReadAsStringAsync()

                    Dim Pushover_Status As String = GetFirstMatchingValueFromJSONResponseString("status", responseBody)

                    If Pushover_Status = "1" Then

                        Dim Pushover_Id As String = GetFirstMatchingValueFromJSONResponseString("id", responseBody)

                        My.Settings.PushoverDeviceID = EncryptionClass.Encrypt(Pushover_Id)
                        My.Settings.PushoverDeviceName = Parms.name ' no need to encrypt
                        My.Settings.Save()

                        LastTimeDataWasReceivedFromPushover = Now

                        ReturnValue = "available"

                    Else

                        If responseBody.Contains("has already been taken") Then
                            '                       
                            If Parms.force = ForceValues.NoForce Then
                                Parms.force = ForceValues.Force
                                SetPushoverDeviceNameAndID_Async(Parms)
                            End If

                        Else

                            ReturnValue = gNotAvailable

                        End If

                    End If

                End Using

            End Using

        Catch ex As Exception

            ReturnValue = gNotAvailable

        End Try

        PushoverDeviceNameAndId = ReturnValue

    End Sub
    Friend Sub DeletePushoverMessages()

        FindHighestPushoverMessageId_Processing = String.Empty
        DeletePushoverMessageID_Processing = String.Empty

        Dim NewThread As Thread = New Thread(AddressOf FindHighestPushoverMessageId_Async)
        NewThread.Start()

        Dim Timeout As DateTime = Now.AddSeconds(10) ' added timeout in v3.5

        While FindHighestPushoverMessageId_Processing = String.Empty

            Thread.Sleep(100)
            If Now > Timeout Then
                PushoverMessageIdToDelete = String.Empty
                Exit While
            End If

        End While

        NewThread = Nothing

        If PushoverMessageIdToDelete = String.Empty Then
            ' there is nothing to delete
        Else

            Dim NewThread2 As Thread = New Thread(AddressOf DeletePushoverMessagesID_Async)
            NewThread2.Start()

            While DeletePushoverMessageID_Processing = String.Empty
                Thread.Sleep(100)
            End While

            If PushoverMessageIdToDelete = String.Empty Then
                Log("Obsolete Pushover notifications for " & My.Settings.PushoverDeviceName & " have been dismissed")
                Log("")
            Else
                Log("Obsolete Pushover notifications for " & My.Settings.PushoverDeviceName & " were not dismissed")
                Log("")
            End If

            NewThread2 = Nothing

        End If

    End Sub

    Friend Event CloseThePushoverWebSocketNow()
    Private Async Sub FindHighestPushoverMessageId_Async()

        FindHighestPushoverMessageId_Processing = String.Empty

        Try

            Dim Status As String = String.Empty
            Dim HighestMessageId As String = String.Empty

            Dim ServerResponse As String = String.Empty

            SendRequest("", "GET", "https://api.pushover.net/1/messages.json" & "?secret=" & EncryptionClass.Decrypt(My.Settings.PushoverSecret) & "&device_id=" & EncryptionClass.Decrypt(My.Settings.PushoverDeviceID), String.Empty, ServerResponse)

            Status = GetFirstMatchingValueFromJSONResponseString("status", ServerResponse)

            If Status = "1" Then

                PushoverMessageIdToDelete = GetLastMatchingValueFromJSONResponseString("id", ServerResponse)

            Else

                Log("Pushover reported a problem with the use of your device! (b)")
                Log("")

                gCriticalPushOverErrorReported = True
                LastTimeDataWasReceivedFromPushover = Now.AddDays(-1)

                RaiseEvent CloseThePushoverWebSocketNow()

            End If

        Catch ex As Exception

        End Try

        FindHighestPushoverMessageId_Processing = "done"

    End Sub

    Friend Function GetLastMatchingValueFromJSONResponseString(ByVal KeyToLookFor As String, ByVal ServerResponse As String) As String

        Dim LastMatchingValue As String = String.Empty

        On Error Resume Next ' changed try catch to on error resume next in 2.4 - to work around problem returning responses involving boolean values

        If ServerResponse.Length > 0 Then

            Dim DictionaryOfJSONResults = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(ServerResponse)

            For Each item In DictionaryOfJSONResults

                Dim ItemKey As String = item.Key
                Dim ItemValue As JArray = item.Value

                If ItemKey = KeyToLookFor Then

                    If ItemValue Is Nothing Then
                        LastMatchingValue = CType(item.Value, String)
                    Else
                        Return item.Value.ToString
                    End If

                ElseIf ItemValue.HasValues Then

                    For Each child In ItemValue

                        For Each ChildProperty As JProperty In child.Children

                            If ChildProperty.Name = KeyToLookFor Then

                                LastMatchingValue = ChildProperty.Value.ToString

                            End If

                        Next

                    Next

                End If

            Next

        End If

        Return LastMatchingValue

    End Function
    Friend Async Sub DeletePushoverMessagesID_Async()

        DeletePushoverMessageID_Processing = String.Empty

        Try

            Dim Status As String = String.Empty

            Dim ServerResponse As String = String.Empty

            Using httpClient = New HttpClient()

                Using request = New HttpRequestMessage(New HttpMethod("POST"), "https://api.pushover.net/1/devices/" & EncryptionClass.Decrypt(My.Settings.PushoverDeviceID) & "/update_highest_message.json")

                    Dim multipartContent = New MultipartFormDataContent()
                    multipartContent.Add(New StringContent(EncryptionClass.Decrypt(My.Settings.PushoverSecret)), "secret")
                    multipartContent.Add(New StringContent(PushoverMessageIdToDelete), "message")
                    request.Content = multipartContent

                    Dim response As HttpResponseMessage = Await httpClient.SendAsync(request)
                    response.EnsureSuccessStatusCode()

                    Dim responseBody As String = Await response.Content.ReadAsStringAsync()

                    Dim Pushover_Status As String = GetFirstMatchingValueFromJSONResponseString("status", responseBody)

                    If Pushover_Status = "1" Then

                        'Logging at this point hangs the system - don't do it
                        PushoverMessageIdToDelete = String.Empty

                    End If

                End Using

            End Using


        Catch ex As Exception

        End Try

        DeletePushoverMessageID_Processing = "done"

    End Sub

    Friend Function GetFirstMatchingValueFromJSONResponseString(ByVal KeyToLookFor As String, ByVal ServerResponse As String) As String

        On Error Resume Next ' changed try catch to on error resume next in 2.4 - to work around problem returning responses involving boolean values

        If ServerResponse.Length > 0 Then

            Dim DictionaryOfJSONResults = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(ServerResponse)

            For Each item In DictionaryOfJSONResults

                Dim ItemKey As String = item.Key
                Dim ItemValue As JArray = item.Value

                If ItemKey = KeyToLookFor Then

                    If ItemValue Is Nothing Then
                        Return CType(item.Value, String)
                    Else
                        Return item.Value.ToString ' changed in v2.5 from ItemValue.ToString
                    End If

                ElseIf ItemValue.HasValues Then

                    For Each child In ItemValue

                        For Each ChildProperty As JProperty In child.Children

                            If ChildProperty.Name = KeyToLookFor Then

                                Return ChildProperty.Value.ToString

                            End If

                        Next

                    Next

                End If

            Next

        End If

    End Function

    ' used for testing 

    'Dim responsebody = "{""errors"":{""name"":[""has already been taken""]},""status"":0,""request"":""9c312212-5178-48d2-9924-f725cc74d842""}"
    'Dim xxx As String = GetUniqueMatchingValueFromJSONResponseString("name", responsebody)

    'Dim responsebody = "{""status"":1,""id"":""uc1dfu3kgqcfuy4n2kcytsy5q8skfm"",""secret"":""sdaeedt58rdycnjwgrbwxiy5n654m5efc2k6bxw97ezy8f24bj5cf23h91qt"",""request"":""46c1d826-ed01-4F8a-8923-b39538efedd2""}"
    'Dim xxx As String = GetUniqueMatchingValueFromJSONResponseString("secret", responsebody)


    'Dim responsebody = "{"messages":[{"id":3,"message":"open the calculator","app":"IFTTT","aid":286,"icon":"ifttt","date":1574008102,"priority":0,"acked":0,"umid":207,"sound":"po","subscription":1},{"id":4,"message":"open the calculator","app":"IFTTT","aid":286,"icon":"ifttt","date":1574008172,"priority":0,"acked":0,"umid":208,"sound":"po","subscription":1},{"id":6,"message":"open the calculator","app":"IFTTT","aid":286,"icon":"ifttt","date":1574008248,"priority":0,"acked":0,"umid":210,"sound":"po","subscription":1}],"user":{"quiet_hours":false,"is_android_licensed":false,"is_ios_licensed":false,"is_desktop_licensed":true,"email":"info@push2run.com","quick_quiet_until":null,"show_team_ad":"1"},"device":{"name":"Push2Run_ROBSPC","encryption_enabled":false,"default_sound":"po","always_use_default_sound":false,"default_high_priority_sound":"po","always_use_default_high_priority_sound":false,"dismissal_sync_enabled":false},"status":1,"request":"7f4d2784-fda7-49d6-9c59-d3e8dc39e242"}
    'Dim xxx As String = GetUniqueMatchingValueFromJSONResponseString("secret", responsebody)

    'xxx = xxx

    'Friend Function GetUniqueMatchingValueFromJSONResponseString(ByVal KeyToLookFor As String, ByVal ServerResponse As String) As String


    '    On Error Resume Next ' changed try catch to on error resume next in 2.4 - to work around problem returning responses involving boolean values

    '    If ServerResponse.Length > 0 Then

    '        Dim DictionaryOfJSONResults = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(ServerResponse)

    '        Dim WorkingResult As String

    '        For Each item As KeyValuePair(Of String, Object) In DictionaryOfJSONResults

    '            Dim ItemKey As String = item.Key
    '            Dim ItemValue As JArray = item.Value

    '            If ItemKey = KeyToLookFor Then

    '                If ItemValue Is Nothing Then
    '                    Return CType(item.Value, String)
    '                Else
    '                    Return item.Value.ToString.TrimStart("{").TrimStart("[").TrimEnd("}").TrimEnd("]").Trim.Trim("""").Trim
    '                End If

    '            Else

    '                If item.Value IsNot Nothing Then
    '                    If item.Value.ToString > String.Empty Then
    '                        WorkingResult = GetFirstMatchingValueFromJSONResponseString(KeyToLookFor, item.Value.ToString)
    '                        If WorkingResult > String.Empty Then
    '                            Return WorkingResult
    '                        End If

    '                    End If
    '                End If

    '            End If

    '        Next

    '    End If

    'End Function

    Friend Function SendRequest(ByVal APIKey As String, ByVal GetOrPost As String, ByVal ServerAddress As String, ByVal Data As String, Optional ByRef ResponseReceived As String = "") As Boolean

        Dim ReturnCode As Boolean = True

        Try

            Dim EncodedData As Byte() = Encoding.ASCII.GetBytes(Data)

            Dim request As System.Net.HttpWebRequest
            request = System.Net.WebRequest.Create(ServerAddress)

            request.ContentType = "application/json"
            request.Credentials = New System.Net.NetworkCredential(APIKey, "")


            If GetOrPost = "GET" Then

                request.Method = "GET"
                request.ContentLength = 0
                ServerAddress &= Data

            Else

                request.Method = "POST"
                If Data.Length > 0 Then
                    request.ContentLength = Data.Length
                    Using requeststream = request.GetRequestStream()
                        requeststream.Write(EncodedData, 0, Data.Length)
                        requeststream.Close()
                    End Using
                End If

            End If

            Dim responseJson As String = vbNull

            Using response As System.Net.HttpWebResponse = request.GetResponse()
                Dim reader = New System.IO.StreamReader(response.GetResponseStream())
                Using (reader)
                    responseJson = reader.ReadToEnd()
                End Using
            End Using

            ResponseReceived = responseJson.ToString

        Catch ex As Exception

            ResponseReceived = ""

            If ServerAddress.ToLower.Contains("pushover") Then
            Else
                LastTimeDataWasReceivedFromPushbullet = Now.AddDays(-1)  '2.5 this only needs to happen when the call is for pushbullet, not pushover
            End If

            ' Log(ex.Message.ToString)  ' causes hang with pushover code

            ReturnCode = False

        End Try

        Return ReturnCode

    End Function


    Friend gImportFileName As String = String.Empty
    Friend Event DoAnImport()
    Friend Sub RaiseAnEventToDoAnImport()
        RaiseEvent DoAnImport()
    End Sub

    Friend Event OptionsClosed()
    Friend Sub RaiseAnEventToCloseOptions()
        RaiseEvent OptionsClosed()
    End Sub

    Friend Event SessionLogClosed()
    Friend Sub RaiseAnEventToCloseSessionLog()
        RaiseEvent SessionLogClosed()
    End Sub

    Friend AboutHelpWindowIsOpen As Boolean = False
    Friend Event AboutClosed()
    Friend Sub RaiseAnEventToShowAboutClosedInSystrayAndMainMenu()
        RaiseEvent AboutClosed()
    End Sub

    Friend Sub KeepHelpOnTop()

        If AboutHelpWindowIsOpen Then
            MakeTopMost(SafeNativeMethods.FindWindow(Nothing, gAboutHelpWindowTitle), My.Settings.AlwaysOnTop)
        End If

    End Sub

    Friend gCapsLock As Boolean
    Friend gNumbLock As Boolean
    Friend gScrollLock As Boolean

    Friend Sub GetLockKeyStates()

        Dim WaitUntilDone As Boolean = Application.Current.Dispatcher.Invoke(Function()
                                                                                 gCapsLock = My.Computer.Keyboard.CapsLock
                                                                                 gNumbLock = My.Computer.Keyboard.NumLock
                                                                                 gScrollLock = My.Computer.Keyboard.ScrollLock
                                                                                 Return True
                                                                             End Function)

        'by using a function (as above) this forces the program to wait for completion, using a sub (as commented out below) allows the code to run in parallel

        'Application.Current.Dispatcher.Invoke(Sub()
        '                                          CAPSLock = Keyboard.GetKeyStates(Key.CapsLock)
        '                                          NumbLock = Keyboard.GetKeyStates(Key.NumLock)
        '                                          ScrollLock = Keyboard.GetKeyStates(Key.Scroll)
        '                                      End Sub)

    End Sub


    Friend Sub CheckInternetToSeeIfANewVersionIsAvailable(ByVal oMe As Object, ByVal Silent As Boolean)

        If My.Settings.CheckForUpdate Then 'v4.4
        Else
            Exit Sub
        End If

        Try

            SeCursor(CursorState.Wait)

            Dim lRunningVersionRelativeToCurrentWebSiteVersion As RunningVersionRelativeToCurrentWebSiteVersion = GetRunningVersionRelativeToCurrentWebSiteVersion()

            SeCursor(CursorState.Normal)

            If lRunningVersionRelativeToCurrentWebSiteVersion = RunningVersionRelativeToCurrentWebSiteVersion.Unknown Then
            Else

                Select Case My.Settings.CheckForUpdateFrequency

                    Case Is = UpdateCheckFrequency.Weekly

                        My.Settings.NextVersionCheckDate = Today.Date.AddDays(7)

                    Case Is = UpdateCheckFrequency.EveryTwoWeeks
                        My.Settings.NextVersionCheckDate = Today.Date.AddDays(14)

                    Case Is = UpdateCheckFrequency.Monthly
                        My.Settings.NextVersionCheckDate = Today.Date.AddMonths(1)

                    Case Else
                        My.Settings.NextVersionCheckDate = Today.Date.AddDays(1)

                End Select

                My.Settings.Save()

                Dim lMostCurrentVersion As String

                With gCurrentVersionAccordingToWebsite
                    lMostCurrentVersion = "v" & .Major & "." & .Minor & "." & .Build & "." & .Revision
                End With

                If Silent Then
                    If My.Settings.SkipUpdateFor = lMostCurrentVersion Then
                        Log("Automatic update check - the most current version is " & lMostCurrentVersion & " however settings indicated this update should be skipped")
                        Log("")
                        Exit Sub
                    End If
                End If


            End If

            Select Case lRunningVersionRelativeToCurrentWebSiteVersion

                Case Is = RunningVersionRelativeToCurrentWebSiteVersion.RunningVersionIsTheSameAsTheCurrentVersion

                    If Silent Then

                        Log("Automatic update check - you are running the most current version of Push2Run")
                        Log("")

                    Else

                        OpenUpgradePromptWindow()

                    End If

                Case Is = RunningVersionRelativeToCurrentWebSiteVersion.RunningVersionIsOlderThanTheCurrentVersion

                    If Silent Then

                        Log("Automatic update check - new version found")
                        Log("")

                    End If

                    OpenUpgradePromptWindow()

                Case Is = RunningVersionRelativeToCurrentWebSiteVersion.RunningVersionIsNewerThanCurrentVersion

                    If Silent Then

                        If gBetaVersionInUse Then
                            Log("Automatic update check - you are running a beta version of Push2Run which is newer than the publicly available version")
                        Else
                            Log("Automatic update check - you are running a newer version of Push2Run than is publicly available")
                        End If

                        Log("")

                    Else

                        OpenUpgradePromptWindow()

                    End If

                Case Is = RunningVersionRelativeToCurrentWebSiteVersion.Unknown

                    Dim Message As String = "Push2Run couldn't determine its most current version." & vbCrLf &
                                            "It is likely that accesses to the Push2Run's website or download host is unavailable right now."

                    If Silent Then

                        Log(Message)
                        Log("")

                    Else

                        Dim Result As MessageBoxResult = TopMostMessageBox(oMe, Message & vbCrLf & "Please try again later.",
                        "Push2Run - Version Check", MessageBoxButton.OK, MessageBoxImage.Information, MessageBoxResult.OK, System.Windows.MessageBoxOptions.None)

                    End If

            End Select

        Catch ex As Exception

        End Try

        SeCursor(CursorState.Normal)

    End Sub

    Private Sub OpenUpgradePromptWindow()

        If gWindowUpgradePromptIsOpen Then

            ' ensures upgrade window is only open once if computer is left running and unattended for several days

        Else

            Log("opening update prompt window")

            Try

                gWindowUpgradePrompt = New WindowUpgradePrompt

                ghCurrentOwner = gCurrentOwner

                gCurrentOwner = gWindowUpgradePrompt.Owner

                gWindowUpgradePrompt.Show()

            Catch ex As Exception
                Log(ex.ToString)
            End Try

        End If

    End Sub



    Friend Function GetRunningVersionRelativeToCurrentWebSiteVersion() As RunningVersionRelativeToCurrentWebSiteVersion
        ' also updates gCurrentVersionAccordingToWebsite

        Dim ReturnValue As RunningVersionRelativeToCurrentWebSiteVersion = RunningVersionRelativeToCurrentWebSiteVersion.Unknown

        Try

            Dim CurrentRunningVersion As String = My.Application.Info.Version.ToString

            Dim myWebClient As System.Net.WebClient = New System.Net.WebClient
            Dim CurrentWebSiteVersionDataFileContents As String = myWebClient.DownloadString(gWebPageVersionCheck)
            myWebClient.Dispose()

            'fix in case retrieved file has vblf and not vbcrlf between first and second row in file
            Dim Entries() As String = Split(CurrentWebSiteVersionDataFileContents, vbCrLf)
            If Entries.Count = 2 Then
            Else
                Entries = Split(CurrentWebSiteVersionDataFileContents, vbLf)
            End If

            Dim CurrentVersion As String = Entries(0).Trim ' the top most number is the current version text file on the web

            Dim FormatedCurrentVerison As String = FormattedNumber(CurrentVersion)
            Dim FormatedRunningVersion As String = FormattedNumber(CurrentRunningVersion)

            If (FormatedRunningVersion = FormatedCurrentVerison) Then
                ReturnValue = RunningVersionRelativeToCurrentWebSiteVersion.RunningVersionIsTheSameAsTheCurrentVersion

            ElseIf (FormatedRunningVersion > FormatedCurrentVerison) Then
                ReturnValue = RunningVersionRelativeToCurrentWebSiteVersion.RunningVersionIsNewerThanCurrentVersion

            Else
                ReturnValue = RunningVersionRelativeToCurrentWebSiteVersion.RunningVersionIsOlderThanTheCurrentVersion

            End If

            Dim Working() As String = FormatedCurrentVerison.Split(".")
            With gCurrentVersionAccordingToWebsite
                .Major = CInt(Working(0))
                .Minor = CInt(Working(1))
                .Build = CInt(Working(2))
                .Revision = CInt(Working(3))
            End With

            Entries = Nothing
            Working = Nothing

        Catch ex As Exception

            With gCurrentVersionAccordingToWebsite
                .Major = 0
                .Minor = 0
                .Build = 0
                .Revision = 0
            End With

        End Try

        Return ReturnValue

    End Function

    <System.Diagnostics.DebuggerStepThrough()> Private Function FormattedNumber(ByVal UnformattedVersionNumber As String) As String

        On Error Resume Next
        Dim Piece() As String = Split(UnformattedVersionNumber, ".")

        Return Piece(0).PadLeft(3, "0"c) & "." & Piece(1).PadLeft(3, "0"c) & "." & Piece(2).PadLeft(3, "0"c) & "." & Piece(3).PadLeft(3, "0"c)

    End Function

    Friend Sub DoEvents() ' replaces application.doevents in wpf

        Try
            Application.Current.Dispatcher.Invoke(DispatcherPriority.Background, New Action(Sub()
                                                                                            End Sub))
        Catch ex As Exception
        End Try

    End Sub


    Friend Function IsDownloadFileForANewerVersion() As Boolean

        Dim ReturnValue As Boolean = False

        Try

            If File.Exists(gAutomaticUpdateLocalDownloadedFileName) Then

                Dim fi As FileInfo = New FileInfo(gAutomaticUpdateLocalDownloadedFileName)

                If fi Is Nothing OrElse fi.Length = 0 Then

                Else

                    Dim DownloadedFileVersion As FileVersion = GetFileVersion(gAutomaticUpdateLocalDownloadedFileName)

                    If DownloadedFileVersion.Major > gCurrentlyRunningVersion.Major Then
                        ReturnValue = True
                    Else
                        If DownloadedFileVersion.Minor > gCurrentlyRunningVersion.Minor Then
                            ReturnValue = True
                        Else
                            If DownloadedFileVersion.Build > gCurrentlyRunningVersion.Build Then
                                ReturnValue = True
                            Else
                                If DownloadedFileVersion.Revision > gCurrentlyRunningVersion.Revision Then
                                    ReturnValue = True
                                End If
                            End If
                        End If

                    End If

                End If

            End If

        Catch ex As Exception
        End Try

        Return ReturnValue

    End Function

    Friend Function IsDownloadFileForTheCurrentlyReleasedVersion() As Boolean

        Dim ReturnValue As Boolean = False

        If gSpecialProcessingForTesting Then

            ReturnValue = True

        Else

            Try

                If File.Exists(gAutomaticUpdateLocalDownloadedFileName) Then

                    Dim fi As FileInfo = New FileInfo(gAutomaticUpdateLocalDownloadedFileName)

                    If fi Is Nothing OrElse fi.Length = 0 Then

                    Else

                        Dim DownloadedFileVersion As FileVersion = GetFileVersion(gAutomaticUpdateLocalDownloadedFileName)

                        ReturnValue = DownloadedFileVersion.Major = gCurrentVersionAccordingToWebsite.Major AndAlso
                                      DownloadedFileVersion.Minor = gCurrentVersionAccordingToWebsite.Minor AndAlso
                                      DownloadedFileVersion.Build = gCurrentVersionAccordingToWebsite.Build AndAlso
                                      DownloadedFileVersion.Revision = gCurrentVersionAccordingToWebsite.Revision

                    End If

                End If

            Catch ex As Exception
            End Try

        End If

        Return ReturnValue

    End Function

    Friend Function GetFileVersion(ByVal filename As String) As FileVersion

        Dim ReturnValue As FileVersion

        Try

            With ReturnValue
                .Major = 0
                .Minor = 0
                .Build = 0
                .Revision = 0
            End With

            If File.Exists(filename) Then

                Dim fi As FileVersionInfo = FileVersionInfo.GetVersionInfo(gAutomaticUpdateLocalDownloadedFileName)
                Dim fv As FileVersion
                With fv
                    .Major = fi.FileMajorPart
                    .Minor = fi.FileMinorPart
                    .Build = fi.FileBuildPart
                    .Revision = fi.FilePrivatePart
                End With

                ReturnValue = fv

            End If

        Catch ex As Exception
        End Try

        Return ReturnValue

    End Function

#Region "Drag And Drop support"

    Friend Function DropIntoCurrentlySelectedRow(ByRef e As DragEventArgs) As Boolean

        Dim ReturnValue = True

        With gCurrentlySelectedRow

            Dim droppedFiles As String() = Nothing

            If e.Data.GetDataPresent("Shell IDList Array") Then

                Try

                    Dim DroppedFileName As String = e.Data.GetData(System.Windows.DataFormats.FileDrop)(0)

                    Dim Extention As String = Path.GetExtension(DroppedFileName).ToUpper

                    Select Case Extention

                        Case Is = ".LNK"

                            'The link file needs to be copied to a place where this program has write permissions - so use the temp path and once done the copied temp link file can be deleted

                            Dim NewPath As String = Path.GetTempPath()

                            Dim FileNameOfDroppedFileName As String = Path.GetFileName(DroppedFileName)
                            Dim TempFileName = Path.Combine(NewPath, FileNameOfDroppedFileName)

                            If File.Exists(TempFileName) Then File.Delete(TempFileName)
                            IO.File.Copy(DroppedFileName, TempFileName)

                            Dim Name As String = String.Empty
                            Dim PathAndFileName As String = String.Empty
                            Dim Description As String = String.Empty
                            Dim WorkingDirectory As String = String.Empty
                            Dim Parms As String = String.Empty
                            Dim KeysToSend As String = String.Empty

                            If GetShortcutInfo(TempFileName, Name, PathAndFileName, Description, WorkingDirectory, Parms) Then

                                If .Description.Length = 0 Then .Description = Name
                                If .Description.Length = 0 Then .Description = Description ' fallback if Name is not available
                                If .ListenFor.Length = 0 Then .ListenFor = "open " & Name & vbCrLf & "start " & Name & vbCrLf & "run " & Name
                                If .Open.Length = 0 Then .Open = PathAndFileName
                                If .StartIn.Length = 0 Then .StartIn = WorkingDirectory
                                If .Parameters.Length = 0 Then .Parameters = Parms

                            End If

                            If File.Exists(TempFileName) Then File.Delete(TempFileName)

                        Case Is = ".EXE", ".BAT", ".VBS"

                            Dim Name As String = Path.GetFileNameWithoutExtension(DroppedFileName)

                            If .Description.Length = 0 Then .Description = Name
                            If .ListenFor.Length = 0 Then .ListenFor = "open " & Name & vbCrLf & "start " & Name & vbCrLf & "run " & Name
                            If .Open.Length = 0 Then .Open = DroppedFileName
                            If .StartIn.Length = 0 Then .StartIn = Path.GetDirectoryName(DroppedFileName)

                        Case Is = ".URL"

                            Dim FullURL As String = GetFileNameOrURLFromLink(DroppedFileName)

                            Dim url = New Uri(FullURL)
                            Dim HostName As String = url.Host.TrimStart("www").TrimStart("WWW").TrimStart(".")

                            If .Description.Length = 0 Then .Description = HostName
                            If .ListenFor.Length = 0 Then .ListenFor = "open " & HostName
                            If .Open.Length = 0 Then .Open = GetFileNameOrURLFromLink(DroppedFileName)

                        Case Is = gPush2RunExtention.ToUpper

                            Dim AllText As String = My.Computer.FileSystem.ReadAllText(DroppedFileName)

                            If AllText.StartsWith("[") Then

                                'v3.2   .p2r export file format

                                gImportFileName = DroppedFileName
                                RaiseEvent DoAnImport()
                                ReturnValue = False

                            Else

                                ' single entry .pr2 file

                                Dim LoadedCard As CardClass = LoadACard(DroppedFileName)

                                .Description = LoadedCard.Description
                                .ListenFor = LoadedCard.ListenFor
                                .Open = LoadedCard.Open
                                .StartIn = LoadedCard.StartDirectory
                                .Parameters = LoadedCard.Parameters
                                .Admin = LoadedCard.StartWithAdminPrivileges
                                .StartingWindowState = LoadedCard.StartingWindowState
                                .KeysToSend = LoadedCard.KeysToSend

                            End If

                        Case Else

                            Dim Filename As String = Path.GetFileName(DroppedFileName)
                            Dim Directory As String = Path.GetDirectoryName(DroppedFileName)
                            Dim Result As StringBuilder = New StringBuilder(1024)

                            Dim rc As IntPtr = FindExecutable(Filename, Directory, Result)

                            If (rc.ToInt32 = 0) OrElse (rc.ToInt32 = 42) Then

                                Dim FilenameWithouExtention As String = Path.GetFileNameWithoutExtension(DroppedFileName)

                                If .Description.Length = 0 Then .Description = FilenameWithouExtention
                                If .ListenFor.Length = 0 Then .ListenFor = "open " & FilenameWithouExtention
                                If .Open.Length = 0 Then .Open = Result.ToString.ToLower
                                If .StartIn.Length = 0 Then .StartIn = Path.GetDirectoryName(Result.ToString.ToLower)
                                If .Parameters.Length = 0 Then .Parameters = """" & DroppedFileName & """" ' v3.4.3.1

                            Else

                                Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "Sorry, I don't know how to handle files with a file type of " & Extention, "Push2Run - Info", MessageBoxButton.OK, MessageBoxImage.Question)

                            End If

                            Result = Nothing

                    End Select

                Catch ex As Exception

                End Try

            End If

        End With

        Return ReturnValue

    End Function


    Private Function GetShortcutInfo(full_name As String, ByRef name As String, ByRef path As String, ByRef descr As String, ByRef working_dir As String, ByRef args As String) As Boolean

        'ref http://csharphelper.com/blog/2012/01/get-information-about-a-windows-shortcut-in-c/

        Dim ReturnCode As Boolean = False

        Try

            name = String.Empty
            path = String.Empty
            descr = String.Empty
            working_dir = String.Empty
            args = String.Empty

            ' Make a Shell object.
            Dim shell As New Shell32.Shell()

            ' Get the shortcut's folder and name.
            Dim shortcut_path As String = full_name.Substring(0, full_name.LastIndexOf("\"))

            Dim shortcut_name As String = full_name.Substring(full_name.LastIndexOf("\") + 1)

            If Not shortcut_name.EndsWith(".lnk") Then shortcut_name &= ".lnk"

            ' Get the shortcut's folder
            Dim shortcut_folder As Shell32.Folder = shell.[NameSpace](shortcut_path)

            ' Get the shortcut's file
            Dim folder_item As Shell32.FolderItem = shortcut_folder.Items().Item(shortcut_name)

            Dim lnk As Shell32.ShellLinkObject = DirectCast(folder_item.GetLink, Shell32.ShellLinkObject)
            name = folder_item.Name
            descr = lnk.Description
            path = lnk.Path

            'bug work around; shell32 returns a path of ... Program Files (x86) ... when it should return ... Program Files ...
            If path.Contains("Program Files (x86)") Then

                If File.Exists(path) Then

                    'no problem program is where it should be

                Else

                    Dim HoldPath As String = path
                    path = path.Replace("Program Files (x86)", "Program Files")

                    If File.Exists(path) Then
                        ' path just needed to be corrected - all good now
                    Else
                        'return path to its original state - but flag the issue to the user
                        path = HoldPath
                        Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "Problem finding program identified inside shortcut", "Push2Run - Warning", MessageBoxButton.OK, MessageBoxImage.Exclamation)
                    End If

                End If

            End If

            working_dir = lnk.WorkingDirectory
            args = lnk.Arguments

            ReturnCode = True

        Catch ex As Exception
            ReturnCode = False
        End Try

        Return ReturnCode

    End Function

    Public Function GetFileNameOrURLFromLink(FileNameOfLink As String) As String

        Dim ReturnValue As String = String.Empty

        Try

            Dim shl = New Shell32.Shell()

            FileNameOfLink = System.IO.Path.GetFullPath(FileNameOfLink)

            Dim dir = shl.[NameSpace](System.IO.Path.GetDirectoryName(FileNameOfLink))

            Dim itm = dir.Items().Item(System.IO.Path.GetFileName(FileNameOfLink))

            Dim lnk As Shell32.ShellLinkObject = DirectCast(itm.GetLink, Shell32.ShellLinkObject)

            ReturnValue = lnk.Target.Path

        Catch ex As Exception

        End Try

        Return ReturnValue

    End Function


    'ref https://stackoverflow.com/questions/9540051/is-an-application-associated-with-a-given-extension/9540278#9540278
    <DllImport("shell32.dll")>
    Private Function FindExecutable(lpFile As String, lpDirectory As String, <Out> lpResult As StringBuilder) As IntPtr
    End Function

#End Region

#Region "Open a webpage"

    Delegate Sub OpenAWebPageDelegate(ByVal WebPage As String)
    Friend Sub OpenAWebPage(ByVal WebPage As String)

        On Error Resume Next
        Process.Start(WebPage)

    End Sub

#End Region

#Region "Control Table"

    Friend Enum ResetEncryptionDecriptionLevel
        Passwords = 0
        Data = 1
    End Enum


    Friend Sub ResetEncryptionAndDecriptionToReadAndWrite(ByVal ResetControl As ResetEncryptionDecriptionLevel)

        '********************************************************************************************************
        'Control1 is used to hold a unique Master_Password for this database (encrypted by default)
        'Control2 is used to hold an indicator yes / no - which as been encrypted by control1 - that says if there is a user password for the boss
        'Control3 is used to hold a user password - encrypted by control1 
        '********************************************************************************************************

        Try

            Dim sSQL As String = "SELECT Control1 , Control3 FROM ControlTable"

            Dim SQLiteConnect As New SQLiteConnection(gSQLiteConnectionString)
            Dim SQLiteCommand As SQLiteCommand = New SQLite.SQLiteCommand(sSQL, SQLiteConnect)

            SQLiteConnect.Open()

            Dim SQLiteDataReader As SQLiteDataReader = SQLiteCommand.ExecuteReader(CommandBehavior.CloseConnection)
            SQLiteDataReader.Read()

            'Note the order of each step below is very important 

            EncryptionClass.ResetEncryptionPassPhrase()
            EncryptionClass.ResetDecryptionPassPhrase()

            EncryptionClass.UpdateEncryptionPassPhrase(EncryptionClass.Decrypt(SQLiteDataReader.GetString(0)))
            EncryptionClass.UpdateDecryptionPassPhrase(EncryptionClass.Decrypt(SQLiteDataReader.GetString(0)))

            If ResetControl = ResetEncryptionDecriptionLevel.Data Then
                EncryptionClass.UpdateEncryptionPassPhrase(EncryptionClass.Decrypt(SQLiteDataReader.GetString(1)))
                EncryptionClass.UpdateDecryptionPassPhrase(EncryptionClass.Decrypt(SQLiteDataReader.GetString(1)))
            End If

            SQLiteDataReader.Close()
            SQLiteConnect.Close()

            SQLiteConnect.Dispose()
            SQLiteCommand.Dispose()
            SQLiteDataReader = Nothing

        Catch ex As Exception

        End Try

    End Sub

#End Region



    Friend Function GetPasswordInPlainText(ByRef Password As String) As Boolean

        '********************************************************************************************************
        'Control1 is used to hold a unique Master_Password for this database (encrypted by default)
        'Control2 is used to hold an indicator yes / no - which as been encrypted by control1 - that says if there is a user password for the boss
        'Control3 is used to hold a user password - encrypted by control1 
        '********************************************************************************************************

        'Exit condition - Decrypt and Encrypt are set up to encode and GUI + Password level

        Dim ReturnCode As String = True

        Try

            Password = String.Empty

            Dim sSQL As String = "SELECT Control1 , Control2 , Control3 FROM ControlTable ;"

            Dim SQLiteConnect As New SQLiteConnection(gSQLiteConnectionString)
            Dim SQLiteCommand As SQLiteCommand = New SQLiteCommand(sSQL, SQLiteConnect)

            SQLiteConnect.Open()

            Dim SQLiteDataReader As SQLiteDataReader = SQLiteCommand.ExecuteReader(CommandBehavior.CloseConnection)

            SQLiteDataReader.Read()

            EncryptionClass.ResetDecryptionPassPhrase()
            Dim Master_Password As String = EncryptionClass.Decrypt(SQLiteDataReader.GetString(0))

            If Master_Password = String.Empty Then
                'someone is screwing around
                ReturnCode = False
            Else
                EncryptionClass.UpdateDecryptionPassPhrase(Master_Password)
            End If

            Dim YesNoFlagBoss As String = EncryptionClass.Decrypt(SQLiteDataReader.GetString(1))
            Dim YesNoFlagWorker As String = EncryptionClass.Decrypt(SQLiteDataReader.GetString(3))

            Password = EncryptionClass.Decrypt(SQLiteDataReader.GetString(2))

            EncryptionClass.ResetDecryptionPassPhrase()
            EncryptionClass.UpdateDecryptionPassPhrase(Master_Password & Password)

            If (YesNoFlagBoss.Length > 3) AndAlso (YesNoFlagBoss.StartsWith("Yes") OrElse YesNoFlagBoss.StartsWith("No ")) Then
            Else
                'someone is screwing around
                ReturnCode = False
            End If

            If YesNoFlagBoss.StartsWith("Yes") AndAlso (Password = "") Then
                Password = GenerateRandomPassword(10)
                ReturnCode = False
            End If

            Master_Password = String.Empty
            YesNoFlagBoss = String.Empty

            SQLiteDataReader.Close()
            SQLiteConnect.Close()

            SQLiteConnect.Dispose()
            SQLiteCommand.Dispose()
            SQLiteDataReader = Nothing

        Catch ex As Exception
            ReturnCode = False
        End Try

        If ReturnCode Then
        Else
            Password = String.Empty
        End If

        Return ReturnCode

    End Function


    Friend Function DoPasswordsMatch(ByRef iEncryptedPassword As String) As Boolean

        '********************************************************************************************************
        'Control1 is used to hold a unique Master_Password for this database (encrypted by default)
        'Control2 is used to hold an indicator yes / no - which as been encrypted by control1 - that says if there is a user password for the boss
        'Control3 is used to hold a user password - encrypted by control1 
        '********************************************************************************************************

        Dim ReturnCode As String

        Try

            Dim sSQL As String = "SELECT Control3 FROM ControlTable ;"

            Dim SQLiteConnect As New SQLiteConnection(gSQLiteConnectionString)
            Dim SQLiteCommand As SQLiteCommand = New SQLiteCommand(sSQL, SQLiteConnect)

            SQLiteConnect.Open()

            Dim SQLiteDataReader As SQLiteDataReader = SQLiteCommand.ExecuteReader(CommandBehavior.CloseConnection)

            SQLiteDataReader.Read()

            ReturnCode = (iEncryptedPassword = SQLiteDataReader.GetString(0))

            SQLiteDataReader.Close()
            SQLiteConnect.Close()

            SQLiteConnect.Dispose()
            SQLiteCommand.Dispose()
            SQLiteDataReader = Nothing

        Catch ex As Exception
            ReturnCode = False
        End Try

        Return ReturnCode

    End Function

    Dim PasswordCheckLock As New Object

    'Friend Function IsAPasswordRequiredForBoss() As Boolean

    '    '********************************************************************************************************
    '    'Control1 is used to hold a unique Master_Password for this database (encrypted by default)
    '    'Control2 is used to hold an indicator yes / no - which as been encrypted by control1 - that says if there is a user password for the boss
    '    'Control3 is used to hold a user password - encrypted by control1 
    '    '********************************************************************************************************

    '    'Exit condition - Decrypt and Encrypt are set up to encode and GUI + Password level

    '    Static InitialCheck As Boolean = False ' added in v4.9.1 to try and avoid a rare deadlock issue

    '    Dim CheckCount As Integer = 0 ' added in v4.9.1 to try and avoid a rare deadlock issue

    '    Dim ReturnCode As Boolean = False ' changed bias in v4.9.1 to false

    '    Dim LineNumber As Integer

    '    While CheckCount < 2

    '        Try

    '            CheckCount += 1

    '            Dim sSQL As String = "SELECT Control1 , Control2 , Control3 FROM ControlTable ;"

    '            Dim SQLiteConnect As New SQLiteConnection(gSQLiteConnectionString)
    '            Dim SQLiteCommand As SQLiteCommand = New SQLiteCommand(sSQL, SQLiteConnect)

    '            LineNumber = 1
    '            SQLiteConnect.Open()

    '            Dim SQLiteDataReader As SQLiteDataReader = SQLiteCommand.ExecuteReader(CommandBehavior.CloseConnection)

    '            LineNumber = 2
    '            SQLiteDataReader.Read()

    '            LineNumber = 3
    '            EncryptionClass.ResetDecryptionPassPhrase()

    '            LineNumber = 4
    '            Dim Master_Password As String = EncryptionClass.Decrypt(SQLiteDataReader.GetString(0))

    '            LineNumber = 5
    '            EncryptionClass.UpdateDecryptionPassPhrase(Master_Password)

    '            LineNumber = 6
    '            Dim YesNoFlag As String = EncryptionClass.Decrypt(SQLiteDataReader.GetString(1))

    '            LineNumber = 7
    '            ReturnCode = (YesNoFlag.Length > 3) AndAlso (YesNoFlag.StartsWith("Yes"))

    '            LineNumber = 8
    '            Dim Password As String = EncryptionClass.Decrypt(SQLiteDataReader.GetString(2))

    '            LineNumber = 9
    '            EncryptionClass.ResetDecryptionPassPhrase()

    '            LineNumber = 10
    '            EncryptionClass.UpdateDecryptionPassPhrase(Master_Password & Password)

    '            LineNumber = 11
    '            Master_Password = String.Empty
    '            YesNoFlag = String.Empty

    '            LineNumber = 12
    '            SQLiteDataReader.Close()
    '            LineNumber = 13
    '            SQLiteConnect.Close()

    '            LineNumber = 14
    '            SQLiteConnect.Dispose()

    '            LineNumber = 15
    '            SQLiteCommand.Dispose()

    '            LineNumber = 16
    '            SQLiteDataReader = Nothing

    '        Catch ex As Exception

    '            ReturnCode = False ' changed bias in v4.9.1 to false

    '            If CheckCount = 1 Then
    '                Log("Password required check fault at line number " & LineNumber.ToString)
    '                Log(ex.Message.ToString)
    '                Log("")
    '                System.Threading.Thread.Sleep(1000)
    '            End If

    '        End Try

    '        InitialCheck = True

    '    End While

    '    InitialCheck = True

    '    Return ReturnCode

    'End Function

    Friend Function IsAPasswordRequiredForBoss() As Boolean
        ' Control1 = Master_Password (encrypted with base phrase)
        ' Control2 = Yes/No flag (encrypted with master)
        ' Control3 = User password (encrypted with master)
        ' Safe default: return False unless a solid “Yes” flag can be confidently decrypted.

        Const LogPrefix As String = "PwdChk:"
        Const MaxAttempts As Integer = 2

        For attempt As Integer = 1 To MaxAttempts
            Try
                Using conn As New SQLiteConnection(gSQLiteConnectionString)
                    conn.Open()
                    Using cmd As New SQLiteCommand("SELECT Control1, Control2, Control3 FROM ControlTable LIMIT 1;", conn)
                        Using rdr As SQLiteDataReader = cmd.ExecuteReader(CommandBehavior.CloseConnection)

                            If Not rdr.Read() Then
                                If attempt = MaxAttempts Then Log(LogPrefix & " control row missing (empty table), default False")
                                Return False
                            End If

                            ' Basic null / empty validation
                            If rdr.IsDBNull(0) OrElse rdr.IsDBNull(1) OrElse rdr.IsDBNull(2) Then
                                If attempt = MaxAttempts Then Log(LogPrefix & " one or more NULL fields, default False")
                                Return False
                            End If

                            Dim encMaster = rdr.GetString(0)
                            Dim encFlag = rdr.GetString(1)
                            Dim encPwd = rdr.GetString(2)

                            If String.IsNullOrWhiteSpace(encMaster) Then
                                If attempt = MaxAttempts Then Log(LogPrefix & " empty master value, default False")
                                Return False
                            End If

                            ' Decrypt master
                            EncryptionClass.ResetDecryptionPassPhrase()
                            Dim master As String
                            Try
                                master = EncryptionClass.Decrypt(encMaster)
                            Catch exDec As Exception
                                If attempt = MaxAttempts Then Log(LogPrefix & " master decrypt fail: " & exDec.Message)
                                Return False
                            End Try
                            If String.IsNullOrEmpty(master) Then
                                If attempt = MaxAttempts Then Log(LogPrefix & " master decrypt empty, default False")
                                Return False
                            End If

                            ' Decrypt flag
                            EncryptionClass.UpdateDecryptionPassPhrase(master)
                            Dim flagPlain As String = String.Empty
                            Try
                                flagPlain = EncryptionClass.Decrypt(encFlag)
                            Catch exFlag As Exception
                                If attempt = MaxAttempts Then Log(LogPrefix & " flag decrypt fail: " & exFlag.Message)
                                Return False
                            End Try

                            ' Decrypt password (even if not needed, to keep encryption state aligned with rest of code)
                            Dim userPwdPlain As String = String.Empty
                            Try
                                userPwdPlain = EncryptionClass.Decrypt(encPwd)
                            Catch exPwd As Exception
                                If attempt = MaxAttempts Then Log(LogPrefix & " password decrypt fail: " & exPwd.Message)
                                Return False
                            End Try

                            ' Final passphrase state per existing pattern
                            EncryptionClass.ResetDecryptionPassPhrase()
                            EncryptionClass.UpdateDecryptionPassPhrase(master & userPwdPlain)

                            Dim result As Boolean = (flagPlain.Length > 3 AndAlso flagPlain.StartsWith("Yes", StringComparison.Ordinal))
                            Return result
                        End Using
                    End Using
                End Using

            Catch ex As Exception
                If attempt = MaxAttempts Then
                    Log(LogPrefix & " exception attempt " & attempt & ": " & ex.Message)
                    Return False
                End If
                ' brief backoff only on first failure
                Thread.Sleep(100)
            End Try
        Next

        Return False
    End Function

    Friend Sub StartupShortCut(ByVal Command As String)

        ' Commands are:  
        ' Add
        ' Delete
        ' Ensure ; ensure will verify the startup link is present and correctly defined 

        Const gThisProgramName As String = "Push2Run"
        Const gReloaderProgramName As String = "Push2RunReloader"

        Try

            Dim ApplicationPathAndFileName As String = System.Environment.CurrentDirectory & "\" & gThisProgramName
            Dim ApplicationPathAndFileNameReloader As String = System.Environment.CurrentDirectory & "\" & gReloaderProgramName

            'If ApplicationPathAndFileName.Contains("Rob Latour") AndAlso ApplicationPathAndFileName.Contains("Debug") Then Exit Sub ' ignore request as this is a debug session

            Dim StartupPathName As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Startup)

            Dim ShortCutFileName_forPush2Run As String = gThisProgramName & ".lnk"
            Dim ShortCutFileName_forPush2RunReloader As String = gReloaderProgramName & ".lnk"

            Dim StartupPathAndShortCutFileName_forPush2Run As String = StartupPathName & "\" & ShortCutFileName_forPush2Run
            Dim StartupPathAndShortCutFileName_forPush2RunReloader As String = StartupPathName & "\" & ShortCutFileName_forPush2RunReloader


            If Command = "Ensure" Then

                'A windows startup link is required
                'read the current startup link and if is pointing to the right program and location then exit out
                'otherwise create/recreate it

                If My.Settings.StartBossOnLogon Then

                    If My.Settings.StartBossAsAdministratorByDefault Then

                        If System.IO.File.Exists(StartupPathAndShortCutFileName_forPush2RunReloader) Then

                            Dim obj As New Object
                            obj = CreateObject("Shell.Application")
                            Dim ShortCutsProgramPathandFileName As String = obj.NameSpace(System.Environment.SpecialFolder.Startup).ParseName(ShortCutFileName_forPush2RunReloader).GetLink.Path.ToString
                            obj = Nothing

                            If ShortCutsProgramPathandFileName = (ApplicationPathAndFileNameReloader & ".exe") Then
                                Exit Sub 'its all good
                            Else
                                Command = "Add" ' there is an short cut there but it points to an incorrect startup location
                            End If

                        Else

                            Command = "Add"

                        End If

                    Else

                        If System.IO.File.Exists(StartupPathAndShortCutFileName_forPush2Run) Then

                            Dim obj As New Object
                            obj = CreateObject("Shell.Application")
                            Dim ShortCutsProgramPathandFileName As String = obj.NameSpace(System.Environment.SpecialFolder.Startup).ParseName(ShortCutFileName_forPush2Run).GetLink.Path.ToString
                            obj = Nothing

                            If ShortCutsProgramPathandFileName = (ApplicationPathAndFileName & ".exe") Then
                                Exit Sub 'its all good
                            Else
                                Command = "Add" ' there is an short cut there but it points to an incorrect startup location
                            End If

                        Else

                            Command = "Add"

                        End If

                    End If

                End If

            End If


            ' Delete any exiting short cuts as they are either a) not needed or b) will be recreated below
            If System.IO.File.Exists(StartupPathAndShortCutFileName_forPush2Run) Then
                System.IO.File.Delete(StartupPathAndShortCutFileName_forPush2Run)
            End If

            If System.IO.File.Exists(StartupPathAndShortCutFileName_forPush2RunReloader) Then
                System.IO.File.Delete(StartupPathAndShortCutFileName_forPush2RunReloader)
            End If


            If Command = "Add" Then

                Dim Shell As Object
                Shell = CreateObject("WScript.Shell")

                Dim MyShortcut As Object

                If My.Settings.StartBossAsAdministratorByDefault Then

                    MyShortcut = Shell.CreateShortcut(StartupPathName & "\" & gReloaderProgramName & ".lnk")
                    MyShortcut.TargetPath = (ApplicationPathAndFileName & ".exe").Replace(gThisProgramName & ".exe", gReloaderProgramName & ".exe")
                    MyShortcut.Arguments = "StartAdmin"

                Else

                    MyShortcut = Shell.CreateShortcut(StartupPathName & "\" & gThisProgramName & ".lnk")
                    MyShortcut.TargetPath = ApplicationPathAndFileName
                    MyShortcut.Arguments = "StartNormal"

                End If

                MyShortcut.WindowStyle = 7 'minimized

                MyShortcut.WorkingDirectory = System.Environment.CurrentDirectory
                MyShortcut.IconLocation = Shell.ExpandEnvironmentStrings(ApplicationPathAndFileName & ".exe" & ", 0")

                MyShortcut.Save()

                MyShortcut = Nothing
                Shell = Nothing

            End If

        Catch ex As Exception
        End Try

    End Sub


    <System.Diagnostics.DebuggerStepThrough()>
    Friend Function GenerateRandomPassword(ByVal PasswordLength As Byte) As String

        Static Dim CharactersToChooseFrom As String = String.Empty

        Randomize()

        If CharactersToChooseFrom = String.Empty Then
            Dim sbWorking As New StringBuilder
            For x As Int32 = 33 To 255 ' v4.9.1 start at 33 to avoid space character
                If (x = 34) OrElse (x = 39) OrElse (x = 145) OrElse (x = 146) OrElse (x = 152) Then
                    'don't use certain quote like characters
                Else
                    sbWorking.Append(Chr(x))
                End If
            Next
            CharactersToChooseFrom = sbWorking.ToString
            sbWorking = Nothing
        End If

        Dim str As New StringBuilder
        Dim CharactersToChooseFromLength As Int32 = CharactersToChooseFrom.Length

        For x As Byte = 1 To PasswordLength
            str.Append(CharactersToChooseFrom.Chars(Rnd() * (CharactersToChooseFromLength - 1)))
        Next

        Return str.ToString

    End Function


#Region "Run Program"

    Friend gIsWindows10OrAbove As Boolean = False
    Friend Sub DetermineIfWindows10OrAbove()

        'see the comments on https://stackoverflow.com/questions/33328739/system-environment-osversion-returns-wrong-version  
        're use of application manifest
        'also https://stackoverflow.com/questions/2819934/detect-windows-version-in-net/8406674

        '    '+------------------------------------------------------------------------------+
        '    '|                    |   PlatformID    |   Major version   |   Minor version   |
        '    '+------------------------------------------------------------------------------+
        '    '| Windows 95         |  Win32Windows   |         4         |          0        |
        '    '| Windows 98         |  Win32Windows   |         4         |         10        |
        '    '| Windows Me         |  Win32Windows   |         4         |         90        |
        '    '| Windows NT 4.0     |  Win32NT        |         4         |          0        |
        '    '| Windows 2000       |  Win32NT        |         5         |          0        |
        '    '| Windows XP         |  Win32NT        |         5         |          1        |
        '    '| Windows 2003       |  Win32NT        |         5         |          2        |
        '    '| Windows Vista      |  Win32NT        |         6         |          0        |
        '    '| Windows 2008       |  Win32NT        |         6         |          0        |
        '    '| Windows 7          |  Win32NT        |         6         |          1        |
        '    '| Windows 2008 R2    |  Win32NT        |         6         |          1        |
        '    '| Windows 8          |  Win32NT        |         6         |          2        |
        '    '| Windows 8.1        |  Win32NT        |         6         |          3        |
        '    '+------------------------------------------------------------------------------+
        '    '| Windows 10         |  Win32NT        |        10         |          0        |
        '    '+------------------------------------------------------------------------------+


        'v4.6
        gIsWindows10OrAbove = (System.Environment.OSVersion.Version.Major >= 10)

    End Sub

    Friend gIsUACOn As Boolean = False
    Friend Sub DetermineIfUACIsOn()

        Try

            Dim myKey As Microsoft.Win32.RegistryKey = My.Computer.Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\Policies\\System", False)

            Dim myKeyValue As Integer = myKey.GetValue("EnableLUA", 0)

            gIsUACOn = (myKeyValue = 1)

        Catch ex As Exception

            gIsUACOn = False

        End Try

        If gIsWindows10OrAbove Then

            IsPromptOnSecureDesktop()
            If gIsPromptOnSecureDesktop Then
            Else
                gIsUACOn = False
            End If

        End If

    End Sub

    Friend gIsPromptOnSecureDesktop As Boolean = False
    Friend Sub IsPromptOnSecureDesktop()

        'For W10 if UAC is on, but there is no prompt, then it is effectively off

        Try

            Dim myKey As Microsoft.Win32.RegistryKey = My.Computer.Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\Policies\\System", False)

            Dim myKeyValue As Integer = myKey.GetValue("PromptOnSecureDesktop", 0)

            gIsPromptOnSecureDesktop = (myKeyValue = 1)

        Catch ex As Exception

            gIsPromptOnSecureDesktop = False

        End Try

    End Sub

    Friend gIsAdministrator As Boolean = False

    Friend Sub DetermineIfProgramIsRunningWithAdministrativePrivileges()


        'v2.1.2         
        'Dim identity = WindowsIdentity.GetCurrent()
        'Dim principal = New WindowsPrincipal(identity)
        'gIsAdministrator = principal.IsInRole(WindowsBuiltInRole.Administrator)


        'v4.6
        Using identity = WindowsIdentity.GetCurrent()
            gIsAdministrator = New WindowsPrincipal(identity).IsInRole(WindowsBuiltInRole.Administrator)
        End Using


    End Sub

    Friend gRunningInASandbox As Boolean = False
    Friend Sub DetermineIfRunningInSandbox()

        ' gRunningInASandbox = System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString.Contains("WDAGUtilityAccount")  ' signifies running in a sandbox

        Using identity = WindowsIdentity.GetCurrent()
            gRunningInASandbox = identity.Name.Contains("WDAGUtilityAccount")
        End Using

    End Sub


    Friend Function RunProgram(ByVal Program As String, ByVal StartInDirectory As String, ByVal Parameters As String, AdministrativePrivilegesRequested As Boolean, ByVal ProcessWindowStyle As ProcessWindowStyle, ByVal DesiredKeysToSend As String) As ActionStatus

        Dim ReturnValue As ActionStatus = ActionStatus.Unknown

        gSendingKeyesIsRequired = (DesiredKeysToSend.Length > 0)

        If Program.Length = 0 Then

            ReturnValue = ActionStatus.NotProcecessNoProgramToRun

        Else

            If Program.ToUpper.Trim = ("ACTIVE WINDOW") Then

                RunActiveWindow(DesiredKeysToSend)

            ElseIf Program.ToUpper.Trim = "DESKTOP" Then

                RunDesktop(DesiredKeysToSend)

            ElseIf Program.ToUpper.Trim = ("MQTT") Then

                If MQTTPublish(Parameters) Then
                    ReturnValue = ActionStatus.Succeeded
                End If

            Else

                Program = Environment.ExpandEnvironmentVariables(Program)
                StartInDirectory = Environment.ExpandEnvironmentVariables(StartInDirectory)
                Parameters = Environment.ExpandEnvironmentVariables(Parameters)

                If (gIsAdministrator AndAlso (Not AdministrativePrivilegesRequested)) Then

                    ReturnValue = RunProgramLowerPrivilege(Program, StartInDirectory, Parameters, ProcessWindowStyle, DesiredKeysToSend)

                Else

                    Dim UACPromptWillBeRequired = ((Not gIsAdministrator) AndAlso (AdministrativePrivilegesRequested))

                    If UACPromptWillBeRequired AndAlso My.Settings.UACLimit Then
                        ReturnValue = ActionStatus.NotProcecessAsAUACPromptWouldBeRequired
                    Else
                        ReturnValue = RunProgramStandard(Program, StartInDirectory, Parameters, AdministrativePrivilegesRequested, ProcessWindowStyle, DesiredKeysToSend)
                    End If

                End If

            End If


        End If

        Return ReturnValue

    End Function

    Friend MQTTClient As MqttClient

    Private Function MQTTPublish(ByVal Parameter As String) As Boolean

        Dim ReturnValue As Boolean = False

        Try

            If My.Settings.UseMQTT Then

                If gMQTTConnectionStatus = ConnectionStatus.Connected Then

                    Dim Topic As String = String.Empty
                    Dim PayLoad As String = String.Empty

                    Dim Result As String = ValidatePublish(Parameter, Topic, PayLoad)

                    Log("MQTT Publish")

                    If (Result = "ok") Then

                        Log("Topic" & vbTab & Topic)
                        Log("Payload" & vbTab & PayLoad)
                        Log("")

                        Publish(Topic, PayLoad).Wait()

                        ReturnValue = (gMQTTPublishingStatus = MQTTPublishingStatus.success)

                    ElseIf Result.StartsWith("Warning") Then

                        Log("Topic" & vbTab & Topic)
                        Log("Payload" & vbTab & PayLoad)
                        Log(Result)
                        Log("")

                        Publish(Topic, PayLoad).Wait()

                        ReturnValue = (gMQTTPublishingStatus = MQTTPublishingStatus.success)

                    Else

                        Log("MQTT Publish")
                        Log(Result)

                    End If

                Else

                    Log("MQTT publishing requested, but the MQTT server Is Not currently connected")
                    Log("request ignored")

                End If

            Else

                Log("MQTT publishing requested, but MQTT Is Not enabled in Push2Run's options")
                Log("request ignored")

            End If

            Log("")

        Catch ex As Exception

        End Try

        Return ReturnValue

    End Function

    Friend Function ValidatePublish(ByVal InputString As String, ByRef Topic As String, ByRef PayLoad As String) As String

        ' Returns 'ok' if InputString can be converted to a Topic and Payload
        ' Note the Topic and Payload are updated ByRef in this Function

        Dim ReturnValue As String = String.Empty

        Try

            InputString = InputString.Trim

            Dim Pos As Integer = InputString.IndexOf(" ")

            Topic = ""
            PayLoad = ""

            If Pos = -1 Then
                Topic = InputString
            Else
                Topic = InputString.Remove(Pos).Trim
                PayLoad = InputString.Remove(0, Pos).Trim
            End If


            If (Topic.Length = 0) AndAlso (PayLoad.Length = 0) Then

                ReturnValue = "The Push2Run card's Parameters field is missing both the MQTT Topic and Payload."
                Exit Try

            End If



            Dim Warning As String = ""

            If Topic.StartsWith("/") Then

                Warning = "The Push2Run card's Parameters field has a MQTT Topic starting with a ('/'). Normally this should be avoided."

            End If


            If Topic.Contains("\") Then

                If Warning > "" Then Warning &= vbCrLf & vbCrLf

                Warning &= "The Push2Run card's Parameters field has a MQTT Topic containing at least one backslash ('\').  Normally only forward slashes ('/') are used."

            End If


            If PayLoad.Length = 0 Then

                If Warning > "" Then Warning &= vbCrLf & vbCrLf

                Warning &= "The Push2Run card's Parameters field doesn't have a MQTT Payload."

            End If


            If Warning > "" Then

                If Warning.Contains(vbCrLf) Then
                    ReturnValue = "Warnings: " & vbCrLf & vbCrLf & Warning
                Else
                    ReturnValue = "Warning: " & vbCrLf & vbCrLf & Warning
                End If

            Else

                ReturnValue = "ok"

            End If

        Catch ex As Exception

            ReturnValue = "Something went wrong when processing the Parameters field" & ex.Message.ToString

        End Try

        Return ReturnValue

    End Function

    Enum MQTTPublishingStatus
        unknown = 0
        success = 1
        failed = 2
    End Enum

    Friend gMQTTPublishingStatus As MQTTPublishingStatus

    Private Async Function Publish(ByVal Topic As String, ByVal Payload As String) As Task

        gMQTTPublishingStatus = MQTTPublishingStatus.unknown

        Try

            Dim messageBuilder As MqttApplicationMessageBuilder = New MqttApplicationMessageBuilder()
            messageBuilder = messageBuilder.WithTopic(Topic)
            messageBuilder = messageBuilder.WithPayload(Encoding.UTF8.GetBytes(Payload))

            Dim message As MqttApplicationMessage = messageBuilder.Build()

            Await MQTTClient.PublishAsync(message).ConfigureAwait(False)
            gMQTTPublishingStatus = MQTTPublishingStatus.success

        Catch ex As Exception
            Log("MQTT publish issue:" & vbCrLf & ex.Message.ToString)
            Log("")

            gMQTTPublishingStatus = MQTTPublishingStatus.failed
        End Try
    End Function


    Friend Function RunDesktop(ByVal DesiredKeysToSend As String) As ActionStatus

        Dim ReturnValue As ActionStatus

        Try

            Dim Handle As IntPtr = SafeNativeMethods.GetDesktopWindow()

            gHandle = Handle

            SafeNativeMethods.SetForegroundWindow(Handle)

            gDesiredKeysToSend = DesiredKeysToSend

            SendKeysToTargetWindow(gDesiredKeysToSend)

            gDesiredKeysToSend = String.Empty

            ReturnValue = ActionStatus.Succeeded

        Catch ex As Exception

            ReturnValue = ActionStatus.Failed

        End Try

        Return ReturnValue

    End Function

    Friend Function RunActiveWindow(ByVal DesiredKeysToSend As String) As ActionStatus

        Dim ReturnValue As ActionStatus

        Try

            Dim Handle As IntPtr = SafeNativeMethods.GetActiveWindow()
            SafeNativeMethods.SetForegroundWindow(Handle)

            gHandle = Handle

            gDesiredKeysToSend = DesiredKeysToSend

            SendKeysToTargetWindow(gDesiredKeysToSend)

            gDesiredKeysToSend = String.Empty

            ReturnValue = ActionStatus.Succeeded

        Catch ex As Exception

            ReturnValue = ActionStatus.Failed

        End Try

        Return ReturnValue

    End Function
    Friend Function RunProgramStandard(ByVal Program As String, ByVal StartInDirectory As String, ByVal Parameters As String, AdministrativePrivilegesRequested As Boolean, ByVal WindowProcessingStyle As ProcessWindowStyle, ByVal DesiredKeysToSend As String) As ActionStatus

        Dim ReturnValue As ActionStatus = ActionStatus.Failed

        Dim myProcess As New Process
        Dim ProcessStarted As Boolean = False

        Try

            With myProcess.StartInfo

                .FileName = Program

                If StartInDirectory.Length > 0 Then .WorkingDirectory = StartInDirectory

                If Parameters.Length > 0 Then .Arguments = Parameters

                If AdministrativePrivilegesRequested Then
                    .Verb = "runas"
                Else
                    .Verb = "open"
                End If

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

            'by setting gDesiredWindowState here, the next window that opens will be set to that window state

            gProcessStartTime = Now ' allows automation to watch and react to a new window
            gDesiredKeysToSend = DesiredKeysToSend

            myProcess.Start()

            ReturnValue = ActionStatus.Succeeded

        Catch ex As Exception

            Log(ex.Message.ToString)

        End Try

        Return ReturnValue

    End Function

#Region "Separating words"

    Friend gSeparatingWords() As String
    Friend gUseSeparatingWords As Boolean
    Friend Sub UpdateGlobalSeparatingWordsArray()

        ' this is done here and only at startup and when the Separating Words table is updated 
        ' result are stored in gSeparatingWords and used when incoming messages are being processed

        ' Each set of Separating words needs to be trimmed, and all sets need to be in an array which has been sorted in reverse order
        ' this so the replacement will have, for example, "and to" replaced before "and", and in this way 

        ' note: by the time they get here Separating words are all in lower case

        gUseSeparatingWords = True

        Dim ws As String = My.Settings.SeparatingWords
        ws = Regex.Replace(ws, "\s+", "") ' remove all white spaces
        ws = ws.Replace(",", "").Trim     ' remove all commas
        If ws.Length = 0 Then
            'if there are no Separating words defined, then override option to use them 
            gUseSeparatingWords = False
        End If

        If gUseSeparatingWords Then

            Dim SeparatingWords() As String = My.Settings.SeparatingWords.Split(",")

            For x As Integer = 0 To SeparatingWords.Count - 1
                SeparatingWords(x) = SeparatingWords(x).Trim
            Next

            Array.Sort(SeparatingWords)

            Array.Reverse(SeparatingWords)

            gSeparatingWords = SeparatingWords

        Else

            gSeparatingWords = {""}

        End If

    End Sub

#End Region

#Region "Windows Automation"

    'ref: https://docs.microsoft.com/en-us/dotnet/api/system.windows.automation.windowpattern.windowopenedevent?view=netframework-4.7.2
    'ref: https://stackoverflow.com/questions/54120120/getting-process-start-with-windowstyle-to-work.
    'ref: https://docs.microsoft.com/en-us/dotnet/api/system.windows.automation.windowpattern.setwindowvisualstate?view=netframework-4.7.2

    'setup ...
    'RegisterForAutomationEvents() 

    Friend gDesiredWindowState As WindowVisualState
    Friend gDesiredKeysToSend As String = String.Empty
    Friend gProcessStartTime As DateTime = Now.AddDays(-1)
    Friend gIgnorAutomationOneTime As Boolean = False

    Private ReadOnly WindowOpenedEvent As AutomationEvent
    Friend Sub RegisterForAutomationEvents()

        Dim eventHandler As AutomationEventHandler = AddressOf OnWindowOpen
        Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent, AutomationElement.RootElement, TreeScope.Children, eventHandler)

    End Sub

    Private LockingObject = New Object

    Private Sub OnWindowOpen(ByVal src As Object, ByVal e As AutomationEventArgs)

        ' can't just test for processor id matching the launched program as the window that opens 
        ' doesn't always have the same pid as was started in the myprocess.start 
        ' (the windows calculator is an example of this)

        ' the following code will look for a new window to be opened in 5 seconds or less of push2run launching a program
        ' if one doesn't get opened in that time frame it will be ignored

        SyncLock LockingObject

            If Now > gProcessStartTime.AddSeconds(5) Then GoTo AllDone

            If gIgnorAutomationOneTime Then
                gIgnorAutomationOneTime = False
                GoTo AllDone
            End If

            gProcessStartTime.AddDays(-1)

            Dim Handle As IntPtr
            Dim OKToSendKeys As Boolean = False

            Try

                If (gDesiredWindowState = WindowVisualState.Minimized) OrElse
                   (gDesiredWindowState = WindowVisualState.Normal) OrElse
                   (gDesiredWindowState = WindowVisualState.Maximized) Then

                    Dim SourceElement As AutomationElement = DirectCast(src, AutomationElement)

                    Handle = SourceElement.Current.NativeWindowHandle

                    'get the window title

                    Dim WindowPattern As WindowPattern = DirectCast(SourceElement.GetCurrentPattern(WindowPattern.Pattern), WindowPattern)

                    Handle = SourceElement.Current.NativeWindowHandle

                    If WindowPattern.WaitForInputIdle(10000) Then
                    Else
                        Exit Try 'object is not responding
                    End If

                    System.Threading.Thread.Sleep(100)

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

                System.Threading.Thread.Sleep(1250)
                SendKeysToTargetWindow(gDesiredKeysToSend)
                gDesiredKeysToSend = String.Empty

            End If

            gDesiredWindowState = Nothing

AllDone:

        End SyncLock

    End Sub

#End Region

#Region "Logging"

    Friend SessionLogIsOpen As Boolean = False
    Friend ContextOfMainWindow As Object  ' this is set in the Main Window as follows: ContextOfMainWindow = Me
    Const MaxLines As Integer = 10000

    'Friend lbLog As New ListBox
    Public SessionLogListViewData As ObservableCollection(Of String) = New ObservableCollection(Of String)()

    Delegate Sub EventHandler()
    Friend Event EventToUpdateTheLogWindowsSessionLog As EventHandler

    Delegate Sub AddToLogDelegate(ByVal Message As String)
    Friend Sub AddToLogWorker(ByVal Message As String)

        SessionLogListViewData.Add(Message)

        If SessionLogListViewData.Count > MaxLines Then SessionLogListViewData.RemoveAt(0)

        If SessionLogIsOpen Then RaiseEvent EventToUpdateTheLogWindowsSessionLog()

        If My.Settings.WriteLogToDisk Then
            My.Computer.FileSystem.WriteAllText(gSessionLogFile, Message & vbCrLf, True)
        End If

        ' no noticeable gain with parallel processing
        '
        'Dim sw As New Stopwatch
        'sw.Start()

        'If My.Settings.WriteLogToDisk Then

        '    Parallel.Invoke(Sub()
        '                        SessionLogListViewData.Add(Message)
        '                        If SessionLogListViewData.Count > MaxLines Then SessionLogListViewData.RemoveAt(0)
        '                        If SessionLogIsOpen Then RaiseEvent EventToUpdateTheLogWindowsSessionLog()
        '                    End Sub,
        '   Sub()
        '       My.Computer.FileSystem.WriteAllText(gSessionLogFile, Message & vbCrLf, True)
        '   End Sub)

        'Else

        '    SessionLogListViewData.Add(Message)
        '    If SessionLogListViewData.Count > MaxLines Then SessionLogListViewData.RemoveAt(0)
        '    If SessionLogIsOpen Then RaiseEvent EventToUpdateTheLogWindowsSessionLog()

        'End If

        'sw.Stop()
        'Console.WriteLine(sw.Elapsed)


    End Sub

    <System.Diagnostics.DebuggerStepThrough()>
    Friend Sub Log(ByVal Message As String)

        Static LastMessage As String = String.Empty

        If Message = String.Empty Then
            If LastMessage = String.Empty Then
                ' don't write two blank lines on after the other
                Exit Sub
            End If
        End If

        LastMessage = Message

        Try

            If Message <> String.Empty Then
                Dim DateStamp As String = Now.ToString("yyyy-MM-dd HH:mm:ss.fff - ")
                Message = DateStamp & Message
            End If

            ContextOfMainWindow.Dispatcher.Invoke(New AddToLogDelegate(AddressOf AddToLogWorker), New Object() {Message})

        Catch ex As Exception

        End Try

    End Sub

#End Region

    Friend Function StringSort(ByVal input As String) As String

        Dim ReturnValue As String = String.Empty

        Dim AllData() As String = input.Split(vbCr)

        'remove unwanted stuff

        For x As Integer = 0 To AllData.Count - 1
            AllData(x) = CleanUpWhiteAndDuplicatedSpaces(AllData(x))
        Next

        AllData = TrimAndMakeLowerCase(AllData) 'v1.5
        AllData = SortAndRemoveDuplicates(AllData)

        For x As Integer = 0 To AllData.Count - 1

            If AllData(x).Trim > String.Empty Then
                ReturnValue &= AllData(x) & vbCrLf
            End If

        Next

        ReturnValue = ReturnValue.Trim

        Return ReturnValue

    End Function

    Friend Function CleanUpWhiteAndDuplicatedSpaces(ByVal input As String) As String

        Dim ReturnValue As String = String.Empty

        Dim arrayofchars = input.ToCharArray()

        'replace all white spaces with a simple space
        For x As Integer = 0 To arrayofchars.Length - 1
            If Char.IsWhiteSpace(arrayofchars(x)) Then
                arrayofchars(x) = " "c
            End If
        Next
        ReturnValue = New String(arrayofchars)

        'remove duplicate spaces
        While ReturnValue <> ReturnValue.Replace("  ", " ")
            ReturnValue = ReturnValue.Replace("  ", " ")
        End While

        Return ReturnValue.Trim

    End Function

    Public Function TrimAndMakeLowerCase(ByVal sender As String()) As String()

        Return Array.ConvertAll(Of String, String)(sender, Function(s) s.Trim.ToLower)

    End Function

    Public Function SortAndRemoveDuplicates(ByVal sender As String()) As String()

        Return (From value In sender Select value Distinct Order By value).ToArray

        ' The following can be used if you want to ignore the case
        'Dim alist As List(Of String) = sender.ToList
        'alist = alist.Distinct(StringComparer.CurrentCultureIgnoreCase).ToList
        'Return alist.ToArray

    End Function


#Region "Database Functions"


    Friend Sub CreateMasterRecord()

        Try

            Dim MasterSwitchRecord As New MyTable1Class

            With MasterSwitchRecord

                .ID = gMasterSwitchID
                .SortOrder = gMasterSwitchSortOrder
                .DesiredStatus = 1
                .WorkingStatus = 1
                .Description = String.Empty 'will be loaded with "Master Switch" later - once all the password stuff has been settled
                .ListenFor = String.Empty
                .Open = String.Empty
                .Parameters = String.Empty
                .StartIn = String.Empty
                .Admin = False
                .StartingWindowState = 0
                .KeysToSend = String.Empty

                InsertARecord(.SortOrder, .DesiredStatus, .WorkingStatus, .Description, .ListenFor, .Open, .Parameters, .StartIn, .Admin, .StartingWindowState, .KeysToSend)

            End With

        Catch ex As Exception
        End Try

    End Sub


    Friend Sub MakeSureMasterRecordIsCorrect()

        Try

            Dim sSQL As String = "UPDATE Table1     " &
              "Set SortOrder           = @SortOrder,     " &
                  "DesiredStatus       = @DesiredStatus, " &
                  "Description         = @Description,   " &
                  "ListenFor           = @ListenFor,     " &
                  "Open                = @Open,          " &
                  "Parameters          = @Parameters,    " &
                  "StartIn             = @StartIn,       " &
                  "Admin               = @Admin,         " &
                  "StartingWindowState = @StartingWindowState, " &
                  "KeysToSend          = @KeysToSend     " &
                  "WHERE ID            = 1 ;"

            Dim SQLiteConnect As New SQLiteConnection(gSQLiteConnectionString)
            Dim SQLiteCommand As SQLiteCommand = New SQLite.SQLiteCommand(sSQL, SQLiteConnect)

            SQLiteConnect.Open()
            SQLiteCommand.Parameters.AddWithValue("SortOrder", 10)

            If gMasterStatus = MonitorStatus.Running Then
                SQLiteCommand.Parameters.AddWithValue("DesiredStatus", MonitorStatus.Running)
            Else
                SQLiteCommand.Parameters.AddWithValue("DesiredStatus", MonitorStatus.Stopped)
            End If

            SQLiteCommand.Parameters.AddWithValue("Description", EncryptionClass.Encrypt(gMasterSwitch))
            SQLiteCommand.Parameters.AddWithValue("ListenFor", "")
            SQLiteCommand.Parameters.AddWithValue("Open", "")
            SQLiteCommand.Parameters.AddWithValue("Parameters", "")
            SQLiteCommand.Parameters.AddWithValue("StartIn", "")
            SQLiteCommand.Parameters.AddWithValue("Admin", False)
            SQLiteCommand.Parameters.AddWithValue("StartingWindowState", 0)
            SQLiteCommand.Parameters.AddWithValue("KeysToSend", "")

            SQLiteCommand.ExecuteNonQuery()

            SQLiteConnect.Close()

            SQLiteConnect.Dispose()
            SQLiteCommand.Dispose()

        Catch ex As Exception
        End Try

    End Sub

    Friend Sub AutoCorrectTable1IfNeeded()

        Try

            'sometimes the sort order gets out of wack; find out if this is the case or not

            Dim IsAutoCorrectRequired As Boolean = False

            Dim sSQL As String = "SELECT ID, SortOrder, Description FROM Table1 ORDER BY ID ASC ;"

            Dim SQLiteConnect As New SQLiteConnection(gSQLiteConnectionString)
            Dim SQLiteCommand As SQLiteCommand = New SQLite.SQLiteCommand(sSQL, SQLiteConnect)

            SQLiteConnect.Open()

            Dim SQLiteDataReader As SQLiteDataReader = SQLiteCommand.ExecuteReader(CommandBehavior.CloseConnection)

            Dim wID, wSortOrder As Integer
            Dim wDescription As String

            Dim CountOfMasterSwitches As Integer = 0

            While SQLiteDataReader.Read()

                wID = SQLiteDataReader.GetInt32(0)
                wSortOrder = SQLiteDataReader.GetInt32(1)

                ' ensure everything is correctly ordered
                If (wID * gGapBetweenSortIDsForDatabaseEntries) <> wSortOrder Then
                    IsAutoCorrectRequired = True
                    Exit While
                End If

                'ensure the master switch appears only once
                wDescription = EncryptionClass.Decrypt(SQLiteDataReader.GetString(2))
                If wDescription = gMasterSwitch Then
                    CountOfMasterSwitches += 1
                    If CountOfMasterSwitches > 1 Then
                        IsAutoCorrectRequired = True
                        Exit While
                    End If
                End If

            End While

            SQLiteDataReader.Close()
            SQLiteConnect.Close()

            SQLiteConnect.Dispose()
            SQLiteCommand.Dispose()
            SQLiteDataReader = Nothing

            If IsAutoCorrectRequired Then

                RebuildTable1_common(My.Settings.SortByDescription)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Friend Sub RebuildTable1_common(ByVal SortByDescription As Boolean)

        Dim BackupFileName As String = gSQLiteFullDatabaseName.Replace(".db3", "_temp_backup_created_" & Now.ToString("yyyy-MM-dd_HH-mm-ss-fff") & ".db3")

        Try

            'Step 1: make a backup
            If File.Exists(BackupFileName) Then File.Delete(BackupFileName)
            File.Copy(gSQLiteFullDatabaseName, BackupFileName)

            'Step 2: read the current database into memory
            Dim Source = New List(Of MyTable1Class)
            Dim sSQL As String = "SELECT * FROM Table1 ORDER BY SortOrder ASC ;"

            Dim SQLiteConnect As New SQLiteConnection(gSQLiteConnectionString)
            Dim SQLiteCommand As SQLiteCommand = New SQLite.SQLiteCommand(sSQL, SQLiteConnect)

            SQLiteConnect.Open()

            Dim SQLiteDataReader As SQLiteDataReader = SQLiteCommand.ExecuteReader(CommandBehavior.CloseConnection)

            While SQLiteDataReader.Read()

                Dim WorkingRecord As New MyTable1Class()

                With WorkingRecord

                    .SortOrder = SQLiteDataReader.GetInt32(DatabaseColumns.SortOrder)
                    .DesiredStatus = SQLiteDataReader.GetInt32(DatabaseColumns.DesiredStatus)
                    .WorkingStatus = SQLiteDataReader.GetInt32(DatabaseColumns.WorkingStatus)
                    .Description = EncryptionClass.Decrypt(SQLiteDataReader.GetString(DatabaseColumns.Description))
                    .ListenFor = EncryptionClass.Decrypt(SQLiteDataReader.GetString(DatabaseColumns.ListenFor))
                    .Open = EncryptionClass.Decrypt(SQLiteDataReader.GetString(DatabaseColumns.Open))
                    .Parameters = EncryptionClass.Decrypt(SQLiteDataReader.GetString(DatabaseColumns.Parameters))
                    .StartIn = EncryptionClass.Decrypt(SQLiteDataReader.GetString(DatabaseColumns.StartIn))
                    .Admin = SQLiteDataReader.GetValue(DatabaseColumns.Admin)
                    .StartingWindowState = SQLiteDataReader.GetValue(DatabaseColumns.StartingWindowState)
                    .KeysToSend = EncryptionClass.Decrypt(SQLiteDataReader.GetString(DatabaseColumns.KeysToSend))

                    Source.Add(WorkingRecord)

                End With

            End While

            'Step 3: read the MasterSwitch for use in Step 5
            Dim MasterSwitchRecord As MyTable1Class = ReadARecord(gMasterSwitchID)

            'Step 4: Remove the current Table1 from the database and reset the auto increment counter to 0
            RunSQL("DELETE FROM Table1 ;")
            RunSQL("UPDATE sqlite_sequence SET seq=0 WHERE name = ""Table1"" ;")

            'Step 5: Reload Table1

            If SortByDescription Then

                ' reorder the Source file
                Source = Source.OrderBy(Function(v) v.Description).ToList()

                'Make sure the Master Switch Record is the first record by inserting it now
                With MasterSwitchRecord
                    InsertARecord(gMasterSwitchSortOrder, .DesiredStatus, .WorkingStatus, .Description, .ListenFor, .Open, .Parameters, .StartIn, .Admin, .StartingWindowState, .KeysToSend)
                End With
                MasterSwitchRecord = Nothing

            End If

            Dim NeededDimensionOfLoadFile As Integer = Source.Count - 1
            Dim LoadFile(NeededDimensionOfLoadFile) As gRowRecord

            Dim x As Integer = 0
            Dim SortOrder As Integer = 0

            If SortByDescription Then
                SortOrder = gGapBetweenSortIDsForDatabaseEntries
            End If

            For Each Entry As MyTable1Class In Source

                With Entry

                    If SortByDescription AndAlso ((.Description = gMasterSwitch) OrElse (.DesiredStatus = StatusValues.NoSwitch)) Then

                        'skip this entry
                        NeededDimensionOfLoadFile -= 1

                    Else

                        ' load the entry into the load file
                        LoadFile(x).SortOrder = SortOrder  ' do not use = x * gGapBetweenSortIDsForDatabaseEntries; it will not work
                        LoadFile(x).DesiredStatus = .DesiredStatus
                        LoadFile(x).WorkingStatus = .WorkingStatus
                        LoadFile(x).Description = .Description
                        LoadFile(x).ListenFor = .ListenFor
                        LoadFile(x).Open = .Open
                        LoadFile(x).Parameters = .Parameters
                        LoadFile(x).StartIn = .StartIn
                        LoadFile(x).Admin = .Admin
                        LoadFile(x).StartingWindowState = .StartingWindowState
                        LoadFile(x).KeysToSend = .KeysToSend

                        SortOrder += gGapBetweenSortIDsForDatabaseEntries
                        x += 1

                    End If

                End With

            Next

            If SortByDescription Then
                ReDim Preserve LoadFile(NeededDimensionOfLoadFile)
            End If

            InsertManyRecords(LoadFile)

            LoadFile = Nothing

            Source = Nothing

            File.Delete(BackupFileName)

        Catch ex As Exception

            Dim Result As MessageBoxResult = TopMostMessageBox(gCurrentOwner, ex.Message.ToString, "Push2Run - Error", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK, System.Windows.MessageBoxOptions.None)

            If File.Exists(BackupFileName) Then
                If File.Exists(gSQLiteFullDatabaseName) Then
                    File.Delete(gSQLiteFullDatabaseName)
                    File.Copy(BackupFileName, gSQLiteFullDatabaseName)
                    File.Delete(BackupFileName)
                End If
            End If

        End Try

    End Sub

#Region "Data Table"

    Friend Enum DatabaseColumns
        ID = 0
        SortOrder = 1
        DesiredStatus = 2
        WorkingStatus = 3
        Description = 4
        ListenFor = 5
        Open = 6
        Parameters = 7
        StartIn = 8
        Admin = 9
        StartingWindowState = 10
        KeysToSend = 11
    End Enum



    Friend Function ReadARecord(ByVal ID As Integer) As MyTable1Class

        Dim ReturnRecord As New MyTable1Class

        Try

            Dim sSQL As String = "SELECT * FROM Table1 WHERE ID=" & ID.ToString.Trim & " ;"

            Dim SQLiteConnect As New SQLiteConnection(gSQLiteConnectionString)
            Dim SQLiteCommand As SQLiteCommand = New SQLite.SQLiteCommand(sSQL, SQLiteConnect)

            SQLiteConnect.Open()

            Dim SQLiteDataReader As SQLiteDataReader = SQLiteCommand.ExecuteReader(CommandBehavior.CloseConnection)

            SQLiteDataReader.Read()

            Dim WorkingRecord As New MyTable1Class()

            With WorkingRecord

                .ID = SQLiteDataReader.GetInt32(DatabaseColumns.ID)
                .SortOrder = SQLiteDataReader.GetInt32(DatabaseColumns.SortOrder)
                .DesiredStatus = SQLiteDataReader.GetInt32(DatabaseColumns.DesiredStatus)
                .WorkingStatus = SQLiteDataReader.GetInt32(DatabaseColumns.WorkingStatus)
                .Description = EncryptionClass.Decrypt(SQLiteDataReader.GetString(DatabaseColumns.Description))
                .ListenFor = EncryptionClass.Decrypt(SQLiteDataReader.GetString(DatabaseColumns.ListenFor))
                .Open = EncryptionClass.Decrypt(SQLiteDataReader.GetString(DatabaseColumns.Open))
                .Parameters = EncryptionClass.Decrypt(SQLiteDataReader.GetString(DatabaseColumns.Parameters))
                .StartIn = EncryptionClass.Decrypt(SQLiteDataReader.GetString(DatabaseColumns.StartIn))
                .Admin = SQLiteDataReader.GetValue(DatabaseColumns.Admin)
                .StartingWindowState = SQLiteDataReader.GetInt32(DatabaseColumns.StartingWindowState)
                .KeysToSend = EncryptionClass.Decrypt(SQLiteDataReader.GetString(DatabaseColumns.KeysToSend))

            End With

            SQLiteDataReader.Close()
            SQLiteConnect.Close()

            SQLiteConnect.Dispose()
            SQLiteCommand.Dispose()
            SQLiteDataReader = Nothing

            ReturnRecord = WorkingRecord

        Catch ex As Exception
            ' Dim Result As MessageBoxResult = TopMostMessageBox(gCurrentOwner, ex.Message.ToString, "Push2Run - Error", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK, System.Windows.MessageBoxOptions.None)
        End Try

        Return ReturnRecord

    End Function

    Friend Function GetLastIDInDatabase() As Integer

        Dim ReturnValue As Integer = 1

        Try

            Dim sSQL As String = "SELECT ID FROM Table1 ORDER BY ID DESC LIMIT 1 ;"

            Dim SQLiteConnect As New SQLiteConnection(gSQLiteConnectionString)
            Dim SQLiteCommand As SQLiteCommand = New SQLite.SQLiteCommand(sSQL, SQLiteConnect)

            SQLiteConnect.Open()

            Dim SQLiteDataReader As SQLiteDataReader = SQLiteCommand.ExecuteReader(CommandBehavior.CloseConnection)

            SQLiteDataReader.Read()

            Dim WorkingRecord As New MyTable1Class()

            With WorkingRecord

                .ID = SQLiteDataReader.GetInt32(DatabaseColumns.ID)

            End With

            SQLiteDataReader.Close()
            SQLiteConnect.Close()

            SQLiteConnect.Dispose()
            SQLiteCommand.Dispose()
            SQLiteDataReader = Nothing

            ReturnValue = WorkingRecord.ID

        Catch ex As Exception
            Dim Result As MessageBoxResult = TopMostMessageBox(gCurrentOwner, ex.Message.ToString, "Push2Run - Error", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK, System.Windows.MessageBoxOptions.None)
        End Try

        Return ReturnValue

    End Function


    Friend Function GetMasterDesiredStatusFromDatabase() As StatusValues


        Dim ReturnCode As StatusValues = StatusValues.SwitchOff

        Try

            Dim sSQL As String = "SELECT DesiredStatus FROM Table1 WHERE ID=1 ;"

            Dim SQLiteConnect As New SQLiteConnection(gSQLiteConnectionString)
            Dim SQLiteCommand As SQLiteCommand = New SQLite.SQLiteCommand(sSQL, SQLiteConnect)

            SQLiteConnect.Open()

            Dim SQLiteDataReader As SQLiteDataReader = SQLiteCommand.ExecuteReader(CommandBehavior.CloseConnection)
            SQLiteDataReader.Read()

            Try

                'v3.2
                'ReturnCode = SQLiteDataReader.GetValue(0)

                Dim Result As Long = SQLiteDataReader.GetValue(0)

                If Result = 1 Then
                    ReturnCode = StatusValues.SwitchOn
                Else
                    ReturnCode = StatusValues.SwitchOff
                End If

            Catch ex As Exception
            End Try


            SQLiteDataReader.Close()
            SQLiteConnect.Close()

            SQLiteConnect.Dispose()
            SQLiteCommand.Dispose()
            SQLiteDataReader = Nothing

        Catch ex As Exception

            Dim Result As MessageBoxResult = TopMostMessageBox(gCurrentOwner, ex.Message.ToString, "Push2Run - Error", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK, System.Windows.MessageBoxOptions.None)

        End Try

        Return ReturnCode

    End Function

    Friend Function CheckDatabaseVersionAndUpgradeItAsNecissary() As Boolean

        'Returns false if there is a problem with the database

        Const Database_from_version_1_0_0_to_version_2_0_5 As String = "ID|SortOrder|DesiredStatus|WorkingStatus|Description|ListenFor|Open|Parameters|StartIn|Admin|"
        Const Database_from_version_2_1_0_to_version_x_x_x As String = "ID|SortOrder|DesiredStatus|WorkingStatus|Description|ListenFor|Open|Parameters|StartIn|Admin|StartingWindowState|KeysToSend|"

        Dim ReturnValue As Boolean = False

        Try

            Dim sSQL As String = "PRAGMA TABLE_INFO(Table1);"

            Dim SQLiteConnect As New SQLiteConnection(gSQLiteConnectionString)
            Dim SQLiteCommand As SQLiteCommand = New SQLite.SQLiteCommand(sSQL, SQLiteConnect)

            SQLiteConnect.Open()

            Dim SQLiteDataReader As SQLiteDataReader = SQLiteCommand.ExecuteReader(CommandBehavior.CloseConnection)

            Dim CurrentColumns As String = String.Empty

            While SQLiteDataReader.Read()
                CurrentColumns &= SQLiteDataReader.GetValue(1) & "|"
            End While

            SQLiteDataReader.Close()
            SQLiteConnect.Close()

            SQLiteConnect.Dispose()
            SQLiteCommand.Dispose()
            SQLiteDataReader = Nothing

            If CurrentColumns = Database_from_version_1_0_0_to_version_2_0_5 Then
                ReturnValue = Upgrade_to_version_2_1()

            ElseIf CurrentColumns = Database_from_version_2_1_0_to_version_x_x_x Then
                ReturnValue = True

            Else
                MsgBox("Unknown database version")

            End If

        Catch ex As Exception
            Dim Result As MessageBoxResult = TopMostMessageBox(gCurrentOwner, ex.Message.ToString, "Push2Run - Error", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK, System.Windows.MessageBoxOptions.None)
        End Try

        Return ReturnValue

    End Function

    Private Function Upgrade_to_version_2_1() As Boolean

        Dim ReturnValue As Boolean = False

        Try

            'backup the current database

            Dim BackupFileName As String = gSQLiteFullDatabaseName.Replace(".db3", "_pre-upgrade_backup_created_" & Now.ToString("yyyy-MM-dd_HH-mm-ss-fff") & ".db3")

            File.Copy(gSQLiteFullDatabaseName, BackupFileName)

            ReturnValue = RunSQLiteSQL("ALTER TABLE Table1 ADD COLUMN StartingWindowState  INTEGER default 0;")
            ReturnValue = RunSQLiteSQL("ALTER TABLE Table1 ADD COLUMN KeysToSend  Text default """";")

            If ReturnValue Then

                Dim SpacerLine As String = String.Empty

                SpacerLine = EncryptionClass.Encrypt(" ")


                ' Update StartingWindowState  values in the column that was just added
                ' StartingWindowState  = 3 is Normal - and is the default for existing entries except where
                ' ID = 1 (the master record) or WorkingStats = 3 (a blank line) 

                ReturnValue = RunSQLiteSQL("UPDATE Table1 SET StartingWindowState = 3 WHERE NOT ( (ID = 1) OR (WorkingStatus = 3) );")

                If ReturnValue Then
                Else
                    MsgBox("Problem in upgrade - alter failed")
                End If

            Else

                MsgBox("Problem in upgrade - update failed")

            End If

        Catch ex As Exception

            MsgBox("Problem in upgrade - " & ex.Message.ToString)

        End Try

        Return ReturnValue

    End Function


    Friend Function RunSQLiteSQL(ByVal myCommandText As String) As Boolean

        Dim ReturnCode As Boolean = False

        Try

            Dim SQLconnect As New SQLiteConnection
            Dim SQLcommand As SQLiteCommand

            SQLconnect.ConnectionString = gSQLiteConnectionString
            SQLconnect.Open()

            SQLcommand = SQLconnect.CreateCommand

            SQLcommand.CommandText = myCommandText

            SQLcommand.ExecuteNonQuery()

            SQLcommand.Dispose()
            SQLconnect.Close()

            ReturnCode = True

        Catch SQLex As SQLiteException
            MsgBox("RunSQLiteSQL 1: " & vbCrLf & myCommandText & vbCrLf & SQLex.ToString)

        Catch ex As Exception
            MsgBox("RunSQLiteSQL 2: " & vbCrLf & gSQLiteConnectionString & vbCrLf & myCommandText & vbCrLf & ex.ToString)

        End Try

        Return ReturnCode

    End Function


    Friend Sub InsertARecord(ByVal iSortOrder As Integer, ByVal iDesiredStatus As Integer, ByVal iWorkingStatus As Integer, ByVal iDescription As String, ByVal iListenFor As String, ByVal iOpen As String, ByVal iParameters As String, ByVal iStartIn As String, ByVal iAdmin As Boolean, ByVal iStartingWindowState As Integer, ByVal iKeysToSend As String)

        Try

            Const sSQL As String = "INSERT INTO Table1 (SortOrder, DesiredStatus, WorkingStatus, Description, ListenFor, Open, Parameters, StartIn, Admin, StartingWindowState, KeysToSend ) " &
                                              "Values(@SortOrder, @DesiredStatus, @WorkingStatus, @Description, @ListenFor, @Open, @Parameters, @StartIn, @Admin, @StartingWindowState, @KeysToSend ) ;"

            Dim SQLiteConnect As New SQLiteConnection(gSQLiteConnectionString)
            Dim SQLiteCommand As SQLiteCommand = New SQLite.SQLiteCommand(sSQL, SQLiteConnect)

            SQLiteConnect.Open()

            SQLiteCommand.Parameters.AddWithValue("SortOrder", iSortOrder)
            SQLiteCommand.Parameters.AddWithValue("DesiredStatus", iDesiredStatus)
            SQLiteCommand.Parameters.AddWithValue("WorkingStatus", iWorkingStatus)
            SQLiteCommand.Parameters.AddWithValue("Description", EncryptionClass.Encrypt(iDescription))
            SQLiteCommand.Parameters.AddWithValue("ListenFor", EncryptionClass.Encrypt(iListenFor))
            SQLiteCommand.Parameters.AddWithValue("Open", EncryptionClass.Encrypt(iOpen))
            SQLiteCommand.Parameters.AddWithValue("Parameters", EncryptionClass.Encrypt(iParameters))
            SQLiteCommand.Parameters.AddWithValue("StartIn", EncryptionClass.Encrypt(iStartIn))
            SQLiteCommand.Parameters.AddWithValue("Admin", iAdmin)
            SQLiteCommand.Parameters.AddWithValue("StartingWindowState", iStartingWindowState)
            SQLiteCommand.Parameters.AddWithValue("KeysToSend", EncryptionClass.Encrypt(iKeysToSend))

            SQLiteCommand.ExecuteNonQuery()

            SQLiteConnect.Close()

            SQLiteConnect.Dispose()
            SQLiteCommand.Dispose()

        Catch ex As Exception
            Dim Result As MessageBoxResult = TopMostMessageBox(gCurrentOwner, ex.Message.ToString, "Push2Run - Error", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK, System.Windows.MessageBoxOptions.None)
        End Try

    End Sub
    Friend Sub InsertManyRecords(ByRef Records As gRowRecord())

        'records are passed note ByRef for performance

        Try

            Const sSQL As String = "INSERT INTO Table1 (SortOrder, DesiredStatus, WorkingStatus, Description, ListenFor, Open, Parameters, StartIn, Admin, StartingWindowState, KeysToSend ) " &
                                                  "Values(@SortOrder, @DesiredStatus, @WorkingStatus, @Description, @ListenFor, @Open, @Parameters, @StartIn, @Admin, @StartingWindowState, @KeysToSend ) ;"

            Static Dim SQLiteConnect As SQLiteConnection
            Static Dim SQLiteCommand As SQLiteCommand

            SQLiteConnect = New SQLiteConnection(gSQLiteConnectionString)
            SQLiteCommand = New SQLite.SQLiteCommand(sSQL, SQLiteConnect)
            SQLiteConnect.Open()

            For Each Record In Records

                With Record

                    SQLiteCommand.Parameters.AddWithValue("SortOrder", .SortOrder)
                    SQLiteCommand.Parameters.AddWithValue("DesiredStatus", .DesiredStatus)
                    SQLiteCommand.Parameters.AddWithValue("WorkingStatus", .WorkingStatus)
                    SQLiteCommand.Parameters.AddWithValue("Description", EncryptionClass.Encrypt(.Description))
                    SQLiteCommand.Parameters.AddWithValue("ListenFor", EncryptionClass.Encrypt(.ListenFor))
                    SQLiteCommand.Parameters.AddWithValue("Open", EncryptionClass.Encrypt(.Open))
                    SQLiteCommand.Parameters.AddWithValue("Parameters", EncryptionClass.Encrypt(.Parameters))
                    SQLiteCommand.Parameters.AddWithValue("StartIn", EncryptionClass.Encrypt(.StartIn))
                    SQLiteCommand.Parameters.AddWithValue("Admin", .Admin)
                    SQLiteCommand.Parameters.AddWithValue("StartingWindowState", .StartingWindowState)
                    SQLiteCommand.Parameters.AddWithValue("KeysToSend", EncryptionClass.Encrypt(.KeysToSend))

                    SQLiteCommand.ExecuteNonQuery()

                End With

            Next

            SQLiteConnect.Close()
            SQLiteConnect.Dispose()
            SQLiteCommand.Dispose()

        Catch ex As Exception
            Dim Result As MessageBoxResult = TopMostMessageBox(gCurrentOwner, ex.Message.ToString, "Push2Run - Error", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK, System.Windows.MessageBoxOptions.None)
        End Try

    End Sub

    <System.Diagnostics.DebuggerStepThrough()>
    <Extension()>
    Friend Function ConvertStartingWindowStateToANumber(ByVal StartingWindowState As String) As Integer

        Dim ReturnValue As Integer

        Select Case StartingWindowState

            Case Is = "Hidden"
                ReturnValue = 1
            Case Is = "Minimized"
                ReturnValue = 2
            Case Is = "Normal"
                ReturnValue = 3
            Case Is = "Maximized"
                ReturnValue = 4
            Case Else
                ReturnValue = 0
        End Select

        Return ReturnValue

    End Function

    <System.Diagnostics.DebuggerStepThrough()>
    <Extension()>
    Friend Function ConvertStartingWindowStateToAString(ByVal StartingWindowState As Integer) As String

        Dim ReturnValue As String

        Select Case StartingWindowState

            Case Is = 1
                ReturnValue = "Hidden"
            Case Is = 2
                ReturnValue = "Minimized"
            Case Is = 3
                ReturnValue = "Normal"
            Case Is = 4
                ReturnValue = "Maximized"
            Case Else
                ReturnValue = ""
        End Select

        Return ReturnValue

    End Function

    <System.Diagnostics.DebuggerStepThrough()>
    <Extension()>
    Friend Function ConvertStartingWindowStateToAProcessWindowStyle(ByVal StartingWindowState As Integer) As ProcessWindowStyle

        Dim ReturnValue As ProcessWindowStyle

        Select Case StartingWindowState

            Case Is = 1
                ReturnValue = ProcessWindowStyle.Hidden
            Case Is = 2
                ReturnValue = ProcessWindowStyle.Minimized
            Case Is = 3
                ReturnValue = ProcessWindowStyle.Normal
            Case Is = 4
                ReturnValue = ProcessWindowStyle.Maximized
            Case Else
                ReturnValue = ProcessWindowStyle.Normal
        End Select

        Return ReturnValue

    End Function


    Friend Sub ChangeARecord(ByVal iID As Integer, ByVal iSortOrder As Integer, ByVal iDesiredStatus As Integer, ByVal iWorkingStatus As Integer, ByVal iDescription As String, ByVal iListenFor As String, ByVal iOpen As String, ByVal iParameters As String, ByVal iStartIn As String, ByVal iAdmin As Boolean, ByVal iStartingWindowState As Integer, ByVal iKeysToSend As String)

        Dim sSQL As String = "UPDATE Table1       " &
            "Set SortOrder            = @SortOrder,     " &
                "DesiredStatus        = @DesiredStatus, " &
                "WorkingStatus        = @WorkingStatus, " &
                "Description          = @Description,   " &
                "ListenFor            = @ListenFor,     " &
                "Open                 = @Open,          " &
                "Parameters           = @Parameters,    " &
                "StartIn              = @StartIn,       " &
                "Admin                = @Admin,         " &
                "StartingWindowState  = @StartingWindowState, " &
                "KeysToSend           = @KeysToSend " &
                "WHERE ID             = @ID ;"

        Try

            Dim SQLiteConnect As New SQLiteConnection(gSQLiteConnectionString)
            Dim SQLiteCommand As SQLiteCommand = New SQLite.SQLiteCommand(sSQL, SQLiteConnect)

            SQLiteConnect.Open()

            SQLiteCommand.Parameters.AddWithValue("SortOrder", iSortOrder)
            SQLiteCommand.Parameters.AddWithValue("DesiredStatus", iDesiredStatus)
            SQLiteCommand.Parameters.AddWithValue("WorkingStatus", iWorkingStatus)
            SQLiteCommand.Parameters.AddWithValue("Description", EncryptionClass.Encrypt(iDescription))
            SQLiteCommand.Parameters.AddWithValue("ListenFor", EncryptionClass.Encrypt(iListenFor))
            SQLiteCommand.Parameters.AddWithValue("Open", EncryptionClass.Encrypt(iOpen))
            SQLiteCommand.Parameters.AddWithValue("Parameters", EncryptionClass.Encrypt(iParameters))
            SQLiteCommand.Parameters.AddWithValue("StartIn", EncryptionClass.Encrypt(iStartIn))
            SQLiteCommand.Parameters.AddWithValue("Admin", iAdmin)
            SQLiteCommand.Parameters.AddWithValue("StartingWindowState", iStartingWindowState)
            SQLiteCommand.Parameters.AddWithValue("KeysToSend", EncryptionClass.Encrypt(iKeysToSend))
            SQLiteCommand.Parameters.AddWithValue("ID", iID)

            SQLiteCommand.ExecuteNonQuery()

            SQLiteConnect.Close()

            SQLiteConnect.Dispose()
            SQLiteCommand.Dispose()

        Catch ex As Exception
            Dim Result As MessageBoxResult = TopMostMessageBox(gCurrentOwner, ex.Message.ToString, "Push2Run - Error", MessageBoxButton.OK, MessageBoxImage.Exclamation, MessageBoxResult.OK, System.Windows.MessageBoxOptions.None)
        End Try

    End Sub


    Friend Sub DeleteARecord(ByVal iID As Integer)

        RunSQL("DELETE FROM Table1 WHERE ID = " & iID.ToString & " ; ")

    End Sub


    Friend Function GetMaxIDFromDatabase() As Integer

        Dim ReturnCode As Integer = -1

        Try

            Dim sSQL As String = "SELECT MAX(ID) FROM Table1 ;"

            Dim SQLiteConnect As New SQLiteConnection(gSQLiteConnectionString)
            Dim SQLiteCommand As SQLiteCommand = New SQLiteCommand(sSQL, SQLiteConnect)

            SQLiteConnect.Open()

            Dim SQLiteDataReader As SQLiteDataReader = SQLiteCommand.ExecuteReader(CommandBehavior.CloseConnection)

            SQLiteDataReader.Read()

            ReturnCode = SQLiteDataReader.GetValue(0)

            SQLiteDataReader.Close()
            SQLiteConnect.Close()

            SQLiteConnect.Dispose()
            SQLiteCommand.Dispose()
            SQLiteDataReader = Nothing

        Catch ex As Exception

            MsgBox("Problem accessing database", MsgBoxStyle.OkOnly, "Push2Run - Warning")

        End Try

        Return ReturnCode

    End Function



    Friend Function GetMaxIDSortOrderFromDatabase() As Integer

        Dim ReturnCode As Integer = -1

        Try

            Dim sSQL As String = "SELECT MAX(SortOrder) FROM Table1 ;"

            Dim SQLiteConnect As New SQLiteConnection(gSQLiteConnectionString)
            Dim SQLiteCommand As SQLiteCommand = New SQLiteCommand(sSQL, SQLiteConnect)

            SQLiteConnect.Open()

            Dim SQLiteDataReader As SQLiteDataReader = SQLiteCommand.ExecuteReader(CommandBehavior.CloseConnection)

            SQLiteDataReader.Read()

            ReturnCode = SQLiteDataReader.GetValue(0)

            SQLiteDataReader.Close()
            SQLiteConnect.Close()

            SQLiteConnect.Dispose()
            SQLiteCommand.Dispose()
            SQLiteDataReader = Nothing

        Catch ex As Exception

            MsgBox("Problem accessing database", MsgBoxStyle.OkOnly, "Push2Run - Warning")

        End Try

        Return ReturnCode


    End Function


    Friend Sub ChangeTheSortOrderOfARecord(ByVal iID As Integer, ByVal iSortOrder As Integer)


        Try

            Dim sSQL As String = "UPDATE Table1   " &
                                 "Set SortOrder = @SortOrder " &
                                 "WHERE ID = @ID ;"

            Dim SQLiteConnect As New SQLiteConnection(gSQLiteConnectionString)
            Dim SQLiteCommand As SQLiteCommand = New SQLite.SQLiteCommand(sSQL, SQLiteConnect)

            SQLiteConnect.Open()

            SQLiteCommand.Parameters.AddWithValue("SortOrder", iSortOrder)
            SQLiteCommand.Parameters.AddWithValue("ID", iID)

            SQLiteCommand.ExecuteNonQuery()

            SQLiteConnect.Close()

            SQLiteConnect.Dispose()
            SQLiteCommand.Dispose()

        Catch ex As Exception
        End Try

    End Sub


    Friend Sub ChangeTheDesiredStatusSwitch(ByVal iID As Integer, ByVal iDesiredStatus As Integer)
        'conversion to sqllite complete

        Dim sSQL As String = "UPDATE Table1   " &
                             "Set DesiredStatus = @DesiredStatus " &
                             "WHERE ID = @ID ;"

        Dim SQLiteConnect As New SQLiteConnection(gSQLiteConnectionString)
        Dim SQLiteCommand As SQLiteCommand = New SQLite.SQLiteCommand(sSQL, SQLiteConnect)

        SQLiteConnect.Open()

        SQLiteCommand.Parameters.AddWithValue("DesiredStatus", iDesiredStatus)
        SQLiteCommand.Parameters.AddWithValue("ID", iID)

        SQLiteCommand.ExecuteNonQuery()

        SQLiteConnect.Close()

        SQLiteConnect.Dispose()
        SQLiteCommand.Dispose()

    End Sub


    Friend Sub ChangeTheWorkingStatusSwitch(ByVal iID As Integer, ByVal iWorkingStatus As Integer)
        'conversion to sqllite complete

        Dim sSQL As String = "UPDATE Table1   " &
                             "Set WorkingStatus = @WorkingStatus " &
                             "WHERE ID = @ID ;"

        Dim SQLiteConnect As New SQLiteConnection(gSQLiteConnectionString)
        Dim SQLiteCommand As SQLiteCommand = New SQLite.SQLiteCommand(sSQL, SQLiteConnect)

        SQLiteConnect.Open()

        SQLiteCommand.Parameters.AddWithValue("WorkingStatus", iWorkingStatus)
        SQLiteCommand.Parameters.AddWithValue("ID", iID)

        SQLiteCommand.ExecuteNonQuery()

        SQLiteConnect.Close()

        SQLiteConnect.Dispose()
        SQLiteCommand.Dispose()

    End Sub




    Friend Function GetEncryptedPassword() As String
        '********************************************************************************************************
        'Control1 is used to hold a unique Master_Password for this database (encrypted by default)
        'Control2 is used to hold an indicator yes / no - which as been encrypted by control1 - that says if there is a user password for the boss
        'Control3 is used to hold a user password - encrypted by control1 
        '********************************************************************************************************

        Dim ReturnCode As String = String.Empty

        Try

            Dim sSQL As String = "SELECT Control3 FROM ControlTable ;"

            Dim SQLiteConnect As New SQLiteConnection(gSQLiteConnectionString)
            Dim SQLiteCommand As SQLiteCommand = New SQLiteCommand(sSQL, SQLiteConnect)

            SQLiteConnect.Open()

            Dim SQLiteDataReader As SQLiteDataReader = SQLiteCommand.ExecuteReader(CommandBehavior.CloseConnection)

            SQLiteDataReader.Read()

            ReturnCode = SQLiteDataReader.GetString(0)

            SQLiteDataReader.Close()
            SQLiteConnect.Close()

            SQLiteConnect.Dispose()
            SQLiteCommand.Dispose()
            SQLiteDataReader = Nothing

        Catch ex As Exception

            MsgBox("Problem accessing database", MsgBoxStyle.OkOnly, "Push2Run - Warning")

        End Try

        Return ReturnCode

    End Function


    Friend Function SetPassword(ByVal NewPassword_PlainText As String) As Boolean

        'Good exit condition - Decrypt and Encrypt are set up to encode and GUI + Password level

        Dim ReturnCode As String = True

        Try

            '********************************************************************************************************
            'Control1 is used to hold a unique Master_Password for this database (encrypted by default)
            'Control2 is used to hold an indicator yes / no - which as been encrypted by control1 - that says if there is a user password for the boss
            'Control3 is used to hold a user password - encrypted by control1 
            '********************************************************************************************************

            '********************************************************************************************************
            'Step 1: retrieve Master_Password and old Password
            'Master_Password is needed to decrypt old Password
            'Old Password is need to decrypt data for encryption transformation to be encrypted under the new password

            Dim sSQL As String = "SELECT Control1 , Control3 FROM ControlTable ;"

            Dim SQLiteConnect As New SQLiteConnection(gSQLiteConnectionString)
            Dim SQLiteCommand As SQLiteCommand = New SQLiteCommand(sSQL, SQLiteConnect)

            SQLiteConnect.Open()

            Dim SQLiteDataReader As SQLiteDataReader = SQLiteCommand.ExecuteReader(CommandBehavior.CloseConnection)

            SQLiteDataReader.Read()

            Dim Master_Password_Encrypted As String = SQLiteDataReader.GetString(0)
            Dim OldPassword_Encrypted As String = SQLiteDataReader.GetString(1)

            SQLiteDataReader.Close()
            SQLiteConnect.Close()

            SQLiteConnect.Dispose()
            SQLiteCommand.Dispose()
            SQLiteDataReader = Nothing

            EncryptionClass.ResetDecryptionPassPhrase()
            Dim Master_Password_PlainText As String = EncryptionClass.Decrypt(Master_Password_Encrypted)

            EncryptionClass.UpdateDecryptionPassPhrase(Master_Password_PlainText)
            Dim OldPassword_PlainText As String = EncryptionClass.Decrypt(OldPassword_Encrypted)

            If (Master_Password_PlainText = "") Then
                ReturnCode = False  'someone is screwing around
            End If

            '********************************************************************************************************
            'Step 2: update new password on database

            Dim NewPassword_Encrypted As String = String.Empty

            EncryptionClass.ResetEncryptionPassPhrase()
            EncryptionClass.UpdateEncryptionPassPhrase(Master_Password_PlainText)

            If NewPassword_PlainText = String.Empty Then
                NewPassword_PlainText = GenerateRandomPassword(10)
                NewPassword_Encrypted = EncryptionClass.Encrypt(NewPassword_PlainText)
            Else
                NewPassword_Encrypted = EncryptionClass.Encrypt(NewPassword_PlainText)
            End If

            Dim sSQL2 As String = "UPDATE ControlTable " &
                                  "SET Control3 = @Control3 ;"

            Dim SQLiteConnect2 As New SQLiteConnection(gSQLiteConnectionString)
            Dim SQLiteCommand2 As SQLiteCommand = New SQLite.SQLiteCommand(sSQL2, SQLiteConnect2)

            SQLiteConnect2.Open()

            SQLiteCommand2.Parameters.AddWithValue("Control3", NewPassword_Encrypted)

            SQLiteCommand2.ExecuteNonQuery()

            SQLiteConnect2.Close()

            SQLiteConnect2.Dispose()
            SQLiteCommand2.Dispose()

            '********************************************************************************************************
            'Step 3: Update the Description, ListenFor, Open, Parameters and StartIn data on Database with new encryption based on new password; other fields are not encrypted and don't need to be reloaded

            EncryptionClass.ResetDecryptionPassPhrase()
            EncryptionClass.UpdateDecryptionPassPhrase(Master_Password_PlainText & OldPassword_PlainText)
            EncryptionClass.ResetEncryptionPassPhrase()
            EncryptionClass.UpdateEncryptionPassPhrase(Master_Password_PlainText & NewPassword_PlainText)

            'Step 3a: Unload
            Dim UnloadTableID(MaxNumberOfEntries) As Integer
            Dim ReEncryptedTableOfDescription(MaxNumberOfEntries) As String
            Dim ReEncryptedTableOfListenFor(MaxNumberOfEntries) As String
            Dim ReEncryptedTableOfOpen(MaxNumberOfEntries) As String
            Dim ReEncryptedTableOfParameters(MaxNumberOfEntries) As String
            Dim ReEncryptedTableOfStartIn(MaxNumberOfEntries) As String
            Dim ReEncryptedTableOfKeysToSend(MaxNumberOfEntries) As String

            Dim UnloadTableIndex As Int32 = 0

            Dim sSQL3 As String = "SELECT ID , Description , ListenFor , Open , Parameters , StartIn, KeysToSend FROM Table1"
            Dim SQLiteConnect3 As New SQLiteConnection(gSQLiteConnectionString)
            Dim SQLiteCommand3 As SQLiteCommand = New SQLite.SQLiteCommand(sSQL3, SQLiteConnect3)

            SQLiteConnect3.Open()

            Dim SQLiteDataReader3 As SQLiteDataReader = SQLiteCommand3.ExecuteReader(CommandBehavior.CloseConnection)

            While SQLiteDataReader3.Read()

                UnloadTableIndex += 1
                UnloadTableID(UnloadTableIndex) = SQLiteDataReader3.GetInt32(0) 'while the unload table id should always = x; using this code will work as a safeguard
                ReEncryptedTableOfDescription(UnloadTableIndex) = EncryptionClass.Encrypt(EncryptionClass.Decrypt(SQLiteDataReader3.GetString(1)))
                ReEncryptedTableOfListenFor(UnloadTableIndex) = EncryptionClass.Encrypt(EncryptionClass.Decrypt(SQLiteDataReader3.GetString(2)))
                ReEncryptedTableOfOpen(UnloadTableIndex) = EncryptionClass.Encrypt(EncryptionClass.Decrypt(SQLiteDataReader3.GetString(3)))
                ReEncryptedTableOfParameters(UnloadTableIndex) = EncryptionClass.Encrypt(EncryptionClass.Decrypt(SQLiteDataReader3.GetString(4)))
                ReEncryptedTableOfStartIn(UnloadTableIndex) = EncryptionClass.Encrypt(EncryptionClass.Decrypt(SQLiteDataReader3.GetString(5)))
                ReEncryptedTableOfKeysToSend(UnloadTableIndex) = EncryptionClass.Encrypt(EncryptionClass.Decrypt(SQLiteDataReader3.GetString(6)))

            End While

            SQLiteDataReader3.Close()
            SQLiteConnect3.Close()

            SQLiteConnect3.Dispose()
            SQLiteCommand3.Dispose()
            SQLiteDataReader3 = Nothing

            'Step 3b Reload Parameters in database with new encryption

            Dim sSQL4 As String = "UPDATE Table1 SET Description = @Description , ListenFor = @ListenFor , Open = @Open , Parameters = @Parameters , Startin = @StartIn , KeysToSend = @KeysToSend WHERE ID = @ID ;"

            Dim SQLiteConnect4 As New SQLiteConnection(gSQLiteConnectionString)
            Dim SQLiteCommand4 As SQLiteCommand = New SQLite.SQLiteCommand(sSQL4, SQLiteConnect4)

            SQLiteConnect4.Open()

            For x As Int32 = 1 To UnloadTableIndex

                SQLiteCommand4.Parameters.AddWithValue("Description", ReEncryptedTableOfDescription(x))
                SQLiteCommand4.Parameters.AddWithValue("ListenFor", ReEncryptedTableOfListenFor(x))
                SQLiteCommand4.Parameters.AddWithValue("Open", ReEncryptedTableOfOpen(x))
                SQLiteCommand4.Parameters.AddWithValue("Parameters", ReEncryptedTableOfParameters(x))
                SQLiteCommand4.Parameters.AddWithValue("StartIn", ReEncryptedTableOfStartIn(x))
                SQLiteCommand4.Parameters.AddWithValue("KeysToSend", ReEncryptedTableOfKeysToSend(x))
                SQLiteCommand4.Parameters.AddWithValue("ID", UnloadTableID(x))
                SQLiteCommand4.ExecuteNonQuery()

                ReEncryptedTableOfParameters(x) = String.Empty
                UnloadTableID(x) = 0

            Next

            SQLiteConnect4.Close()

            SQLiteConnect4.Dispose()
            SQLiteCommand4.Dispose()

            ReDim UnloadTableID(0)
            ReDim ReEncryptedTableOfParameters(0)

            UnloadTableID = Nothing
            ReEncryptedTableOfParameters = Nothing

            '********************************************************************************************************
            'Step 5: Clean up 

            EncryptionClass.ResetDecryptionPassPhrase()
            EncryptionClass.UpdateDecryptionPassPhrase(Master_Password_PlainText & NewPassword_PlainText)
            'EncryptDecryptClass.ResetEncryptionPassPhrase()  ' done above
            'EncryptDecryptClass.UpdateEncryptionPassPhrase(Master_Password_PlainText & NewPassword_PlainText) 'done above

        Catch ex As Exception
            ReturnCode = False
        End Try

        Return ReturnCode

    End Function


    Friend Function SetPasswordIsRequireFlag(ByVal IsRequired As Boolean) As Boolean

        'Good exit condition - Decrypt and Encrypt are set up to encode and GUI + Password level

        Dim ReturnCode As String = True

        Try

            '********************************************************************************************************
            'Control1 is used to hold a unique Master_Password for this database (encrypted by default)
            'Control2 is used to hold an indicator yes / no - which as been encrypted by control1 - that says if there is a user password for the boss
            'Control3 is used to hold a user password - encrypted by control1 
            '********************************************************************************************************

            Dim sSQL As String = "SELECT Control1 , Control3 FROM ControlTable ;"

            Dim SQLiteConnect As New SQLiteConnection(gSQLiteConnectionString)
            Dim SQLiteCommand As SQLiteCommand = New SQLiteCommand(sSQL, SQLiteConnect)

            SQLiteConnect.Open()

            Dim SQLiteDataReader As SQLiteDataReader = SQLiteCommand.ExecuteReader(CommandBehavior.CloseConnection)

            SQLiteDataReader.Read()

            Dim Master_Password_Encrypted As String = SQLiteDataReader.GetString(0)
            Dim User_Password_Encrypted As String = SQLiteDataReader.GetString(1)

            SQLiteDataReader.Close()
            SQLiteConnect.Close()

            SQLiteConnect.Dispose()
            SQLiteCommand.Dispose()
            SQLiteDataReader = Nothing

            '********************************************************************************************************

            Dim WhichControlValueToUpdate As String = String.Empty

            EncryptionClass.ResetDecryptionPassPhrase()
            Dim Master_Password_PlainText As String = EncryptionClass.Decrypt(Master_Password_Encrypted)

            EncryptionClass.ResetEncryptionPassPhrase()
            EncryptionClass.UpdateEncryptionPassPhrase(Master_Password_PlainText)

            Dim NewYesNOFlag_Encrypted As String = String.Empty
            If IsRequired Then
                NewYesNOFlag_Encrypted = EncryptionClass.Encrypt("Yes" & GenerateRandomPassword(7))
            Else
                NewYesNOFlag_Encrypted = EncryptionClass.Encrypt("No " & GenerateRandomPassword(7))
            End If

            '********************************************************************************************************

            Dim sSQL2 As String = "UPDATE ControlTable SET Control2 = @ControlValue ;"

            Dim SQLiteConnect2 As New SQLiteConnection(gSQLiteConnectionString)
            Dim SQLiteCommand2 As SQLiteCommand = New SQLite.SQLiteCommand(sSQL2, SQLiteConnect2)

            SQLiteConnect2.Open()

            SQLiteCommand2.Parameters.AddWithValue("ControlValue", NewYesNOFlag_Encrypted)

            SQLiteCommand2.ExecuteNonQuery()

            SQLiteConnect2.Close()

            SQLiteConnect2.Dispose()
            SQLiteCommand2.Dispose()

            '********************************************************************************************************
            'Step 5: Clean up

            ResetEncryptionAndDecriptionToReadAndWrite(ResetEncryptionDecriptionLevel.Data)

            Master_Password_PlainText = String.Empty
            User_Password_Encrypted = String.Empty

        Catch ex As Exception
            ReturnCode = False
        End Try

        Return ReturnCode

    End Function

    <System.Diagnostics.DebuggerStepThrough()>
    Friend Function RunSQL(ByVal sSQL As String, Optional ByVal DisplayErrors As Boolean = True) As Boolean


        Dim ReturnCode As Boolean = True

        Try

            Dim SQLiteConnect As New SQLiteConnection(gSQLiteConnectionString)
            Dim SQLiteCommand As SQLiteCommand = New SQLiteCommand(sSQL, SQLiteConnect)

            SQLiteConnect.Open()

            SQLiteCommand.ExecuteNonQuery()

            SQLiteConnect.Close()

            SQLiteConnect.Dispose()
            SQLiteCommand.Dispose()

        Catch ex1 As SQLiteException
            ReturnCode = False
            If DisplayErrors Then
                ' MsgBox("Run SQL Error 1:" & vbCrLf & sSQL & vbCrLf & ex1.ToString)
            End If

        Catch ex As Exception
            ReturnCode = False
            If DisplayErrors Then
                ' MsgBox("Run SQL error 2:" & vbCrLf & sSQL & vbCrLf & ex.ToString)
            End If
        End Try

        Return ReturnCode

    End Function

#End Region

#End Region

#End Region

#Region "Top Most stuff"

    <System.Diagnostics.DebuggerStepThrough()>
    Friend Sub MakeTopMost(ByVal hwnd As IntPtr, ByVal TopMost As Boolean)

        Const HWND_TOPMOST As Integer = -1
        Const HWND_NOTOPMOST As Integer = -2
        Const SWP_NOMOVE As Integer = &H2
        Const SWP_NOSIZE As Integer = &H1
        Const TOPMOST_FLAGS As Integer = SWP_NOMOVE Or SWP_NOSIZE

        Dim Dummy As Boolean

        If TopMost Then
            Dummy = SafeNativeMethods.SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS)
        Else
            Dummy = SafeNativeMethods.SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS)
        End If

    End Sub

    Friend Function TopMostMessageBox(ByVal Owner As Window, ByVal Message As String, ByVal Title As String, Button As MessageBoxButton, MessageBoxImage As MessageBoxImage, Optional ByVal MessageBoxResult As MessageBoxResult = MessageBoxButton.OK, Optional ByVal Options As System.Windows.MessageBoxOptions = System.Windows.MessageBoxOptions.None) As MessageBoxResult

        'keeps the message box on top

        If (Application.Current.Dispatcher.CheckAccess()) Then

            Return MessageBox.Show(Owner, Message, Title, Button, MessageBoxImage, MessageBoxResult, Options)

        Else

            Return Application.Current.Dispatcher.Invoke(Function()
                                                             Return MessageBox.Show(Owner, Message, Title, Button, MessageBoxImage, MessageBoxResult, Options)
                                                         End Function)
        End If

    End Function

#End Region


    Friend Sub CheckForUpdatePrompt_DontWaitForAResponse(ByVal Owner As Window, ByVal Message As String, ByVal Title As String, Button As MessageBoxButton, MessageBoxImage As MessageBoxImage, Optional ByVal MessageBoxResult As MessageBoxResult = MessageBoxButton.OK, Optional ByVal Options As System.Windows.MessageBoxOptions = System.Windows.MessageBoxOptions.None)

        'keeps the message box on top; can respond any time - specific to check version logic

        If (Application.Current.Dispatcher.CheckAccess()) Then

            System.Threading.ThreadPool.QueueUserWorkItem(AddressOf BackgroundTask, New Object() {Owner, Message, Title, Button, MessageBoxImage, MessageBoxResult, Options})

        Else

            Application.Current.Dispatcher.Invoke(Function()
                                                      Return System.Threading.ThreadPool.QueueUserWorkItem(AddressOf BackgroundTask, New Object() {Owner, Message, Title, Button, MessageBoxImage, MessageBoxResult, Options})
                                                  End Function)
        End If

    End Sub


    Friend Sub BackgroundTask(ByVal state As Object)

        Try

            Dim Result As MessageBoxResult

            Dim array() As Object = CType(state, Object())
            Dim Owner As Object = array(0)
            Dim Message As String = array(1)
            Dim Title As String = array(2)
            Dim Button As MessageBoxButton = array(3)
            Dim MessageBoxImage As MessageBoxImage = array(4)
            Dim MessageBoxResult As MessageBoxResult = array(5)
            Dim Options As System.Windows.MessageBoxOptions = array(6)

            Result = MessageBox.Show(Message, Title, Button, MessageBoxImage, MessageBoxResult, Options)

            If Result = MessageBoxResult.Yes Then
                Owner.Dispatcher.Invoke(New OpenAWebPageDelegate(AddressOf OpenAWebPage), New Object() {gWebPageDownload})
            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Friend Function IsValidFileName(ByVal fn As String) As Boolean

        If fn.Length = 0 Then
            Return False
        Else
            Return (fn.ToCharArray.Intersect(IO.Path.GetInvalidFileNameChars).Count = 0)
        End If

    End Function

    Friend Function IsValidPathName(ByVal p As String) As Boolean

        If p.Length = 0 Then
            Return False
        Else

            If p.EndsWith("\\") Then
                Return False
            Else
                Return (p.ToCharArray.Intersect(IO.Path.GetInvalidPathChars).Count = 0)
            End If

        End If

    End Function


#Region "Setting the Cursor"
    Friend Enum CursorState
        Normal = 0
        Wait = 1
    End Enum
    Friend Sub SeCursor(ByVal CursorState As CursorState)

        If CursorState = CursorState.Normal Then
            Application.Current.Dispatcher.Invoke(New SetNormalCursorCallback(AddressOf SetNormalCursor))
        Else
            Application.Current.Dispatcher.Invoke(New SetWaitCursorCallback(AddressOf SetWaitCursor))
        End If

    End Sub

    Delegate Sub SetWaitCursorCallback()
    Private Sub SetWaitCursor()
        Mouse.OverrideCursor = Cursors.Wait
    End Sub

    Delegate Sub SetNormalCursorCallback()
    Private Sub SetNormalCursor()
        Mouse.OverrideCursor = Nothing
    End Sub

#End Region

#Region "Toast"

    'ref: https//docs.microsoft.com/en-us/windows/apps/design/shell/tiles-And-notifications/send-local-toast?tabs=uwp#step-2-send-a-toast
    'ref: https://docs.microsoft.com/en-us/windows/apps/design/shell/tiles-and-notifications/adaptive-interactive-toasts?tabs=builder-syntax

    'Not currently used - kept in case it is needed later; may need to fiddle with this further (second example works, first might work)
    'example 1: Dim datauri As New Uri("file:///" + Path.GetFullPath("logo.png"), UriKind.Absolute)
    'example 2: Dim datauri As New Uri("file:///E:\Documents\VBNet\Push2Run\Push2Run\Resources\logo.png", UriKind.Absolute)
    'Toast.AddAppLogoOverride(datauri) ', ToastGenericAppLogoCrop.Circle)

    Friend Sub ToastNotification(ByVal Line1 As String, ByVal Line2 As String, ByVal Line3 As String, ByVal ToastTime As DateTime)

        'call primitive based on the options set in the Notification Options window

        ToastNotificationPrimative(Line1, Line2, Line3, My.Settings.ShowNotificationResult, My.Settings.ShowNotificationSource, ToastTime)

    End Sub

    Friend Sub ToastNotificationForNetorkEvents(ByVal Line1 As String, ByVal Line2 As String, ByVal ToastTime As DateTime)

        If gInitialStartupUnderway OrElse gShutdownUnderway Then

        Else

            If My.Settings.IncludeDisconnectAndReconnect Then

                ToastNotificationPrimative(Line1, Line2, "", True, False, ToastTime)

            End If

        End If

    End Sub

    Friend Sub ToastNotificationPrimative(ByVal Line1 As String, ByVal Line2 As String, ByVal Line3 As String, ByVal Line2Needed As Boolean, ByVal Line3Needed As Boolean, ToastTime As DateTime)

        Application.Current.Dispatcher.Invoke(Sub()

                                                  Try

                                                      'Log("Toast: " & Line1 & " : " & If(Line2Needed, Line2 & " : ", "_ : ") & If(Line3Needed, Line3, "_"))

                                                      Dim Toast As ToastContentBuilder = New ToastContentBuilder()

                                                      Toast.AddText(Line1)

                                                      If Line2Needed Then Toast.AddText(Line2)

                                                      If Line3Needed Then Toast.AddText(Line3)

                                                      With ToastTime
                                                          Toast.AddCustomTimeStamp(New DateTime(.Year, .Month, .Day, .Hour, .Minute, .Second, DateTimeKind.Local))
                                                      End With

                                                      Toast.Show()

                                                      Toast = Nothing

                                                  Catch ex As Exception

                                                      Log("Toast failed:" & vbCrLf & ex.Message)

                                                  End Try

                                              End Sub)

    End Sub

#End Region

End Module
