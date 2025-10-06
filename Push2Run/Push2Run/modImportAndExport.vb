Imports System.Data
Imports System.Data.SQLite
Imports System.IO
'Imports System.Windows.Forms
Imports Microsoft.Win32
Imports Newtonsoft.Json.Linq

Module modImportAndExport
    Friend Sub ExportDatabase()

        Dim gExportPathName As String = String.Empty
        Dim gExportFileName = String.Empty

        Try

            If System.IO.Directory.Exists(My.Settings.ImportExportDirectory) Then
                gExportPathName = My.Settings.ImportExportDirectory
            Else
                gExportPathName = Environment.SpecialFolder.MyDocuments
            End If

            Dim folderBrowserDialog1 As New Forms.FolderBrowserDialog
            folderBrowserDialog1.Description = "Select a folder to export the Push2Run database into." & vbCrLf & "Please note, the exported file's contents will not be encrypted."
            folderBrowserDialog1.SelectedPath = gExportPathName

            If folderBrowserDialog1.ShowDialog() = Forms.DialogResult.OK Then

                gExportFileName = folderBrowserDialog1.SelectedPath & "\Push2Run_database_extract_" & Now.ToString("yyyy-MM-dd_HH-mm-ss") & ".p2r"

                Log("Starting export process")
                Log("")
                Log("Exported database to " & gExportFileName)
                Log("")

                My.Settings.ImportExportDirectory = folderBrowserDialog1.SelectedPath
                My.Settings.Save()

            Else

                Exit Sub

            End If

        Catch ex As Exception

            Log("Something went wrong with the export process")
            Log(ex.Message.ToString)
            Log("Export terminated")
            Exit Sub

        End Try

        Try

            If File.Exists(gSQLiteFullDatabaseName) Then
            Else
                Exit Sub
            End If

            If File.Exists(gExportFileName) Then
                File.Delete(gExportFileName)
            End If

            Dim WorkingID As Integer
            Dim Description As String
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
            Dim SQLiteCommand2 As SQLiteCommand = New System.Data.SQLite.SQLiteCommand(sSQL2, SQLiteConnect2)

            SQLiteConnect2.Open()

            Dim SQLiteDataReader2 As SQLiteDataReader = SQLiteCommand2.ExecuteReader(CommandBehavior.CloseConnection)

            My.Computer.FileSystem.WriteAllText(gExportFileName, "[" & vbCrLf, False)

            Dim FirstEntry As Boolean = True

            SQLiteDataReader2.Read() ' gets rid of the master switch

            While SQLiteDataReader2.Read()

                DesiredStatus = StatusValues.SwitchOff
                Description = String.Empty
                ListenFor = String.Empty
                Open = String.Empty
                Parameters = String.Empty
                StartIn = String.Empty
                Admin = False
                StartingWindowState = 0
                KeysToSend = String.Empty

                WorkingID = SQLiteDataReader2.GetInt32(DatabaseColumns.ID)
                Description = SQLiteDataReader2.GetString(DatabaseColumns.Description)
                DesiredStatus = SQLiteDataReader2.GetInt32(DatabaseColumns.DesiredStatus)
                ListenFor = SQLiteDataReader2.GetString(DatabaseColumns.ListenFor)  ' Defer decrypting info until the code below (2 separate places) when we know it will be needed
                Open = SQLiteDataReader2.GetString(DatabaseColumns.Open) ' Defer decrypting info until the code below (2 separate places) when we know it will be needed
                Parameters = SQLiteDataReader2.GetString(DatabaseColumns.Parameters) ' Defer decrypting info until the code below (2 separate places) when we know it will be needed
                StartIn = SQLiteDataReader2.GetString(DatabaseColumns.StartIn) ' Defer decrypting info until the code below (2 separate places) when we know it will be needed
                Admin = SQLiteDataReader2.GetValue(DatabaseColumns.Admin)
                StartingWindowState = SQLiteDataReader2.GetInt32(DatabaseColumns.StartingWindowState)
                KeysToSend = SQLiteDataReader2.GetString(DatabaseColumns.KeysToSend) ' Defer decrypting info until the code below (2 separate places) when we know it will be needed

                If FirstEntry OrElse (DesiredStatus = 3) Then
                    FirstEntry = False
                Else
                    My.Computer.FileSystem.WriteAllText(gExportFileName, "," & vbCrLf, True)
                End If

                If DesiredStatus = 3 Then

                    ' ignore blank lines

                Else

                    Dim JsonRecord As JObject = New JObject(
                                              New JProperty("Description", EncryptionClass.Decrypt(Description)),
                                              New JProperty("ListenFor", EncryptionClass.Decrypt(ListenFor)),
                                              New JProperty("Open", EncryptionClass.Decrypt(Open)),
                                              New JProperty("Parameters", EncryptionClass.Decrypt(Parameters)),
                                              New JProperty("StartIn", EncryptionClass.Decrypt(StartIn)),
                                              New JProperty("Admin", Admin),
                                              New JProperty("StartingWindowState", StartingWindowState),
                                              New JProperty("KeysToSend", EncryptionClass.Decrypt(KeysToSend)))

                    My.Computer.FileSystem.WriteAllText(gExportFileName, JsonRecord.ToString(), True)

                End If

            End While

            My.Computer.FileSystem.WriteAllText(gExportFileName, vbCrLf & "]", True)

            SQLiteDataReader2.Close()
            SQLiteConnect2.Close()
            SQLiteConnect2.Dispose()
            SQLiteCommand2.Dispose()
            SQLiteDataReader2 = Nothing

            Dim SafteyCheck As String = File.ReadAllText(gExportFileName)
            If SafteyCheck.StartsWith("[" & vbCrLf & "," & vbCrLf & "{") Then

                While SafteyCheck.StartsWith("[" & vbCrLf & "," & vbCrLf & "{")
                    SafteyCheck = SafteyCheck.Replace("[" & vbCrLf & "," & vbCrLf & "{", "[" & vbCrLf & "{")
                End While

                My.Computer.FileSystem.WriteAllText(gExportFileName, SafteyCheck, False)

            End If

            Log("Please note, the exported file's contents are not encrypted")
            Log("")

            Log("Exporting completed!")
            Log("")

            Beep()
            Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "Export complete!" & vbCrLf & vbCrLf & "The exported file's name is:" & vbCrLf & gExportFileName & vbCrLf & vbCrLf & "Please note, the exported file's contents are not encrypted.", "Push2Run - Export complete", MessageBoxButton.OK, MessageBoxImage.Information)

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub


    <System.Diagnostics.DebuggerStepThrough()>
    Sub ClearWorkingRecord(ByRef workingRecord As gImportExportRecord)

        With workingRecord
            .Description = String.Empty
            .ListenFor = String.Empty
            .Open = String.Empty
            .Parameters = String.Empty
            .StartIn = String.Empty
            .Admin = False
            .StartingWindowState = 0
            .KeysToSend = String.Empty
        End With

    End Sub
    Friend Sub ImportDatabase(ByVal passedFilename As String)

        Dim ImportFileName As String = String.Empty
        Dim EntriesInImportFile = 0
        Dim ImportedCards As Integer = 0

        Dim ImportFailed As Boolean = False

        Try

            If passedFilename.Length > 0 Then

                ImportFileName = passedFilename

            Else

                Dim openFileDialog1 = New OpenFileDialog

                If System.IO.Directory.Exists(My.Settings.ImportExportDirectory) Then
                    openFileDialog1.InitialDirectory = My.Settings.ImportExportDirectory
                Else
                    openFileDialog1.InitialDirectory = "c:\"
                End If

                openFileDialog1.Filter = "Push2Run files (*.p2r)|*.p2r|All files (*.*)|*.*"
                openFileDialog1.FilterIndex = 1

                If openFileDialog1.ShowDialog() Then

                    ImportFileName = openFileDialog1.FileName

                    My.Settings.ImportExportDirectory = System.IO.Path.GetDirectoryName(ImportFileName)
                    My.Settings.Save()

                Else
                    Exit Sub
                End If

            End If

            Log("Importing started")
            Log("")
            Log("Importing to database from " & ImportFileName)
            Log("")

        Catch ex As Exception

            Log("Something went wrong with the import process")
            Log(ex.Message.ToString)
            Log("Import terminated")
            Exit Sub

        End Try


        Try

            Dim json As String = File.ReadAllText(ImportFileName)
            Dim AllEntries As JArray = JArray.Parse(json)

            Dim workingRecord As gImportExportRecord

            LoadAllEntriesFromDatabase()

            For Each Entry As JObject In AllEntries

                EntriesInImportFile += 1

                ClearWorkingRecord(workingRecord)

                Dim data As List(Of JToken) = Entry.Children().ToList

                For Each item As JProperty In data

                    item.CreateReader()

                    With workingRecord

                        Select Case item.Name

                            Case "Description", "Descrption" ' there was a typo in the word Description in releases up to and including v4.9; both spellings are supported here for backwards compatibility
                                .Description = item.Value

                            Case "ListenFor"
                                .ListenFor = item.Value

                            Case "Open"
                                .Open = item.Value

                            Case "Parameters"
                                .Parameters = item.Value

                            Case "StartIn"
                                .StartIn = item.Value

                            Case "Admin"
                                .Admin = (item.Value.ToString.Trim.ToLower = "true")

                            Case "StartingWindowState"
                                .StartingWindowState = item.Value

                            Case "KeysToSend"
                                .KeysToSend = item.Value

                        End Select

                    End With

                Next

                Log("Importing entry with description: """ & workingRecord.Description & """")

                If AddImportedRecordIfNotAlreadyInDatabase(workingRecord) Then
                    ImportedCards += 1
                End If

            Next

        Catch ex As Exception
            ImportFailed = True
            Log("")
            Log("Import failed - most likely the Import file is in the wrong format")
        End Try

        If ImportFailed Then
        Else
            Log("")
            Log("Importing complete!")
        End If

        Log("")

        Beep()

        If My.Settings.ImportConfirmation Then

            If ImportFailed Then
                Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "Import process failed!" & vbCrLf & vbCrLf & "No new cards were imported." & vbCrLf & vbCrLf & "For more information please view the Session log.", "Push2Run - Import failed", MessageBoxButton.OK, MessageBoxImage.Information)
            ElseIf ImportedCards = 0 Then
                Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "Import process complete!" & vbCrLf & vbCrLf & "No new cards were imported." & vbCrLf & vbCrLf & "For more information please view the Session log.", "Push2Run - Import complete", MessageBoxButton.OK, MessageBoxImage.Information)
            ElseIf ImportedCards = 1 Then
                Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "Import process complete!" & vbCrLf & vbCrLf & "1 new card was imported." & vbCrLf & vbCrLf & "For more information please view the Session log.", "Push2Run - Import complete", MessageBoxButton.OK, MessageBoxImage.Information)
            Else
                Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "Import process complete!" & vbCrLf & vbCrLf & ImportedCards & " new cards were imported." & vbCrLf & vbCrLf & "For more information please view the Session log.", "Push2Run - Import complete", MessageBoxButton.OK, MessageBoxImage.Information)
            End If

        End If


    End Sub

    Private Function AddImportedRecordIfNotAlreadyInDatabase(ByVal importedRecord As gImportExportRecord) As Boolean

        Dim ReturnValue As Boolean = False

        With importedRecord

            Try

                If ImportedEntryAlreadyExists(importedRecord) Then
                    Log("    similar card already in the database")
                    Exit Try
                End If

                If (.Description.Trim.Length = 0) AndAlso (.ListenFor.Trim.Length = 0) AndAlso (.Open.Trim.Length = 0) AndAlso (.Parameters.Trim.Length = 0) AndAlso (.StartIn.Length = 0) AndAlso (Not .Admin) AndAlso (.KeysToSend.Length = 0) Then
                    Log("    blank card ignored")
                    Exit Try
                End If

                Try

                    .StartIn = .StartIn.Trim

                    If .StartIn.Trim.Length > 0 Then
                        If Directory.Exists(Environment.ExpandEnvironmentVariables(.StartIn.Trim)) Then
                            ' continue on 
                        Else
                            Log("    warning - Start directory does not exist.")
                        End If
                    End If

                    'Validate .Open field

                    If .Open.ToUpper.Trim = "DESKTOP" Then

                        .Open = "Desktop"
                        If .Admin Then
                            Log("    info - Desktop cannot be open with admin privileges; setting 'Admin' to false.")
                            .Admin = False
                        End If

                    End If

                    If (.Open.ToUpper.Trim = "ACTIVE WINDOW") Then

                        .Open = "Active Window"
                        If .Admin Then
                            Log("    info - the Active Window cannot be open with admin privileges; setting 'Admin' to false.")
                            .Admin = False
                        End If

                    End If

                    Dim ExecutableFileExtenstions() As String = {"", ".exe", ".bat", ".vbs", ".ps1"}

                    'search for executable program as entered or with one of the above extensions
                    'in the case the file name as entered would include the full path
                    For Each Filetype As String In ExecutableFileExtenstions
                        If File.Exists(Environment.ExpandEnvironmentVariables(.Open & Filetype)) Then
                            'all is good
                            Exit Try
                        End If
                    Next

                    Dim Filename As String = Path.GetFileName(Environment.ExpandEnvironmentVariables(.Open))

                    ' look for program in the Push2Run Card's .Start Directory
                    If .StartIn.Length > 0 Then
                        For Each Filetype As String In ExecutableFileExtenstions
                            If File.Exists(FolderAndFileCombine(Environment.ExpandEnvironmentVariables(.StartIn), Environment.ExpandEnvironmentVariables(Filename & Filetype))) Then
                                'all good
                                Exit Try
                            End If
                        Next
                    End If

                    'search for executable program in the system path as entered or with one of the above extensions
                    'in the case the file name as entered would not include the full path
                    For Each Filetype As String In ExecutableFileExtenstions
                        If SearchForAFileInTheSystemPath(Environment.ExpandEnvironmentVariables(Filename & Filetype)) Then
                            'all is good
                            Exit Try
                        End If
                    Next

                    'search for executable program in sysnative as appropriate
                    Dim SysPath As String = Environment.SystemDirectory.ToUpper
                    Dim AlternateSystem32FolderName As String = SysPath.Replace("WINDOWS\SYSTEM32", "Windows\Sysnative")

                    For Each Filetype As String In ExecutableFileExtenstions
                        If File.Exists(Environment.ExpandEnvironmentVariables(AlternateSystem32FolderName & "\" & Filename & Filetype)) Then
                            Log("    info - the adjusting .Open and .Startup entries to suite your system.")
                            .Open = Filename & Filetype
                            .StartIn = AlternateSystem32FolderName
                        End If
                    Next

                    'search the registry

                    If .StartIn.Trim = String.Empty Then

                        Dim TestProgramName As String = Environment.ExpandEnvironmentVariables(.Open)

                        If TestProgramName.ToLower.EndsWith(".exe") Then
                        Else
                            TestProgramName &= ".exe"
                        End If

                        Dim MatchFound As Boolean = False

                        Try

                            Dim registry_key(1) As String
                            registry_key(0) = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths"
                            registry_key(1) = "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths"

                            For x As Integer = 0 To 1

                                Using key As Microsoft.Win32.RegistryKey = Registry.LocalMachine.OpenSubKey(registry_key(x))
                                    For Each subkey_name As String In key.GetSubKeyNames()

                                        If subkey_name.ToLower = TestProgramName.ToLower Then

                                            Using subkey As RegistryKey = key.OpenSubKey(subkey_name)

                                                Dim name = DirectCast(subkey.GetValue("Path"), String)
                                                If Not String.IsNullOrEmpty(name) Then
                                                    .StartIn = name
                                                    MatchFound = True
                                                    Exit For

                                                End If

                                            End Using

                                        End If

                                    Next

                                End Using

                                If MatchFound Then Exit For

                            Next

                        Catch ex As Exception
                        End Try

                        If MatchFound Then
                            If (Environment.ExpandEnvironmentVariables(.Open) <> Environment.ExpandEnvironmentVariables(.StartIn & "\" & TestProgramName)) Then
                                .Open = .StartIn & "\" & TestProgramName
                                Log("    info - the adjusting .Open and .Startup entries to suite your system.")
                            End If
                        Else

                            If (.Open.ToUpper.Trim = "DESKTOP") OrElse (.Open.ToUpper.Trim = "ACTIVE WINDOW") OrElse (.Open.ToUpper.Trim = "HTTP") OrElse (.Open.ToUpper.Trim = "WWW") Then
                            Else
                                Log("    warning - the program to open couldn't be found.")
                            End If

                        End If

                    End If

                Catch ex As Exception
                End Try

                Try

                    Dim SortOrder As Integer = GetMaxIDSortOrderFromDatabase()
                    SortOrder += gGapBetweenSortIDsForDatabaseEntries

                    Dim DesiredSwitchState As StatusValues
                    If My.Settings.ImportOnByDefault Then
                        DesiredSwitchState = SwitchValues.SwitchOn
                    Else
                        DesiredSwitchState = SwitchValues.SwitchOff
                    End If

                    Dim WorkingSwitchState As StatusValues = DesiredSwitchState
                    If gMasterStatus = MonitorStatus.Running Then
                        WorkingSwitchState = DesiredSwitchState
                    Else
                        WorkingSwitchState = StatusValues.SwitchOff
                    End If

                    If My.Settings.ImportTag Then
                        If .Description.Trim.EndsWith("(Imported)") Then
                        Else
                            .Description &= " (Imported)"
                        End If
                    End If

                    With importedRecord
                        InsertARecord(SortOrder, DesiredSwitchState, WorkingSwitchState, .Description, .ListenFor, .Open, .Parameters, .StartIn, .Admin, .StartingWindowState, .KeysToSend)
                    End With

                    ReturnValue = True

                Catch ex As Exception
                End Try

            Catch ex As Exception

            End Try

        End With

        If ReturnValue Then
            Log("    card imported.")
        Else
            Log("    card was not imported.")
        End If

        Return ReturnValue

    End Function

    Private Function ImportedEntryAlreadyExists(ByVal importedEntry As gImportExportRecord) As Boolean

        Dim MatchFound As Boolean = False

        For Each entry In AllEntriesTable

            With importedEntry

                ' don't check for Description, as its the rest of the stuff the determines if the record is a duplicate or not

                If (entry.ListenFor = .ListenFor) AndAlso
                   (entry.Open = .Open) AndAlso
                   (entry.Parameters = .Parameters) AndAlso
                   (entry.StartIn = .StartIn) AndAlso
                   (entry.Admin = .Admin) AndAlso
                   (entry.StartingWindowState = .StartingWindowState) AndAlso
                   (entry.KeysToSend = .KeysToSend) Then

                    MatchFound = True

                    Exit For

                End If

            End With

        Next

        Return MatchFound

    End Function


End Module
