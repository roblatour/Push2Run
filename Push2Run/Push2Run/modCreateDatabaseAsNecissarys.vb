Imports System.Data.SQLite
Imports System.IO
Imports System.Reflection
Imports System.Threading

Module modCreateDatabaseAsNecissary

    Private Structure gTable1RecordStructure
        Friend ID As Long
        Friend SortOrder As Integer
        Friend DesiredStatus As Integer
        Friend WorkingStatus As Integer
        Friend Description As String
        Friend ListenFor As String
        Friend Open As String
        Friend Parameters As String
        Friend KeysToSend As String
    End Structure

    Private OriginalTable1Entries(MaxNumberOfEntries) As gTable1RecordStructure

    
    Friend Sub CreateDatabaseAsNecissary()

        If File.Exists(gSQLiteFullDatabaseName) Then

        Else

            EncryptionClass.ResetDecryptionPassPhrase()
            EncryptionClass.ResetEncryptionPassPhrase()
            EncryptionClass.UpdateDecryptionPassPhrase("LoadFromAccess" & My.Computer.Name.ToString)

            CreateANewDatabase()

        End If

    End Sub

    
    Private Sub CreateANewDatabase()

        If CreateSQLiteDatabase() Then
            DoEvents() 'v4.6
            System.Threading.Thread.Sleep(2000)
        Else
            MsgBox("Create a new database failed",, "Push2Run - Critical error")
        End If

    End Sub

    
    Private Function CreateSQLiteDatabase() As Boolean

        Dim CreateWasOK As Boolean = False

        Try
            DoEvents() 'v4.6
            If CreateSQLDataBaseFile() Then
                DoEvents() 'v4.6
                System.Threading.Thread.Sleep(200)
                If CreateControlTable() Then
                    DoEvents() 'v4.6
                    System.Threading.Thread.Sleep(200)
                    If CreateTable1() Then
                        DoEvents() 'v4.6
                        System.Threading.Thread.Sleep(200)
                        CreateMasterRecord()
                        DoEvents() 'v4.6
                        System.Threading.Thread.Sleep(200)
                        If LoadSQLPassword() Then
                            DoEvents() 'v4.6
                            System.Threading.Thread.Sleep(200)
                            InsertARecord(gGapBetweenSortIDsForDatabaseEntries, StatusValues.NoSwitch, StatusValues.NoSwitch, "", "", "", "", "", False, 0, "")
                            DoEvents() 'v4.6
                            System.Threading.Thread.Sleep(200)

                            'v4.6
                            If gRunningInASandbox Then
                            Else
                                'only install the calculator example if not running in the sandbox, as the calculator is not usable in the sandbox
                                InsertARecord(gGapBetweenSortIDsForDatabaseEntries * 2, StatusValues.SwitchOn, StatusValues.SwitchOn, "Calculator", "open the calculator" & vbCrLf & "start the calculator", "calc", "", "C:\Windows\System32\", False, 3, "") 'v2.0.4 added startup directory for calculator
                                DoEvents() 'v4.6
                                System.Threading.Thread.Sleep(200)
                            End If

                            InsertARecord(gGapBetweenSortIDsForDatabaseEntries * 2, StatusValues.SwitchOn, StatusValues.SwitchOn, "Open Notepad and do some typing", "open notepad" & vbCrLf & "start notepad", "notepad", "", "C:\Windows\System32", False, 3, "Hello World!") 'v4.6
                            DoEvents() 'v4.6
                            System.Threading.Thread.Sleep(200)

                        End If
                        CreateWasOK = True
                        'gDatabaseWasCreated = True
                    End If
                End If
            End If

        Catch ex As Exception

            MsgBox("Create SQL database failed" & vbCrLf & ex.Message,, "Push2Run - Critical error")

        End Try

        Return CreateWasOK

    End Function

    
    Private Function CreateSQLDataBaseFile() As Boolean

        Dim CreateWasOK As Boolean = False

        Try

            CreateFullDirectory(Path.GetDirectoryName(gSQLiteFullDatabaseName))
            System.Threading.Thread.Sleep(200) 'v4.6
            If File.Exists(gSQLiteFullDatabaseName) Then File.Delete(gSQLiteFullDatabaseName)
            System.Threading.Thread.Sleep(200)
            SQLiteConnection.CreateFile(gSQLiteFullDatabaseName.Replace("\", "\\"))
            ' System.Threading.Thread.Sleep(200)

            System.Threading.Thread.Sleep(1000)
            If IsFileAvailable(gSQLiteFullDatabaseName, 5000, 100) Then
                CreateWasOK = True
            Else
                CreateWasOK = False
                MsgBox("Problem creating " & gSQLiteFullDatabaseName)

            End If

        Catch SQLex As SQLiteException

            Trace.WriteLine("CreateSQLDataBaseFile SQL Error" & vbCrLf & SQLex.ToString)
            MsgBox("Create SQL database file failed." & vbCrLf & SQLex.Message,, "Push2Run - Critical error")

        Catch ex As Exception

            Trace.WriteLine("CreateSQLDataBaseFile Exception" & vbCrLf & ex.ToString)
            MsgBox("Create SQL database file failed!" & vbCrLf & ex.Message,, "Push2Run - Critical error")

        End Try

        Return CreateWasOK

    End Function

    
    Private Sub CreateFullDirectory(ByVal DirectoryName As String)

        On Error Resume Next

        If DirectoryName.Length < 4 Then Exit Sub ' min length = 3 ; example "c:\"

        Dim ws As String = Microsoft.VisualBasic.Left(DirectoryName, 3)
        For x As Int32 = 4 To Len(DirectoryName)
            If Mid(DirectoryName, x, 1) = "\" Then
                If Not Directory.Exists(ws) Then Directory.CreateDirectory(ws)
            End If
            ws &= Mid(DirectoryName, x, 1)
        Next

        If DirectoryName.EndsWith("\") Then
        Else

            Directory.CreateDirectory(ws)
            Thread.Sleep(200) 'v4.6

        End If

    End Sub

    
    Private Function CreateControlTable() As Boolean

        Dim ReturnCode As Boolean

        '********************************************************************************************************
        'Control1 is used to hold a unique Master_Password for this database (encrypted by default)
        'Control2 is used to hold an indicator yes / no - which as been encrypted by control1 - that says if there is a user password for the boss
        'Control3 is used to hold a user password - encrypted by control1 
        '********************************************************************************************************

        ReturnCode = RunSQLiteSQL("CREATE TABLE ControlTable ( Control1 TEXT PRIMARY KEY , Control2 TEXT , Control3 TEXT );")

        Return ReturnCode

    End Function

    
    Friend Function CreateTable1() As Boolean

        Dim ReturnCode As Boolean

        ReturnCode = RunSQLiteSQL("CREATE TABLE Table1 (" &
                           "ID INTEGER PRIMARY KEY AUTOINCREMENT, " &
                           "SortOrder INTEGER, " &
                           "DesiredStatus INTEGER, " &
                           "WorkingStatus INTEGER, " &
                           "Description TEXT, " &
                           "ListenFor TEXT, " &
                           "Open TEXT, " &
                           "Parameters TEXT, " &
                           "StartIn TEXT, " &
                           "Admin Boolean, " &
                           "StartingWindowState INTEGER, " &
                           "KeysToSend TEXT " &
                           ");"
                          )

        Return ReturnCode

    End Function

    
    Private Function LoadSQLPassword() As Boolean

        Dim LoadWasOK As Boolean = False

        'A Master_Password is generated at the current time of the creation of the sql database and used as part of the password formula

        'This Master_Password is encrypted and stored in the database in a control record

        'The unencrypted Master_Password is also added to the password as stored in the settings file

        'This ties the password to the database together

        'exit condition: 
        '   decrypt with: "LoadFromAccess" & My.Computer.Name.ToString
        '   encrypt with Master_Password + Password

        Try

            'Step 1 update database with the control records

            Dim SQLconnect As New SQLiteConnection
            Dim SQLcommand As SQLiteCommand

            SQLconnect.ConnectionString = gSQLiteConnectionString
            SQLconnect.Open()

            SQLcommand = SQLconnect.CreateCommand

            SQLcommand.CommandText =
                "INSERT INTO ControlTable (  Control1 ,  Control2 ,  Control3 ) " &
                               "   VALUES ( @Control1 , @Control2 , @Control3 ) ;"

            '********************************************************************************************************
            'Control1 is used to hold a unique Master_Password for this database (encrypted by default)
            'Control2 is used to hold an indicator yes / no - which as been encrypted by control1 - that says if there is a user password for the boss
            'Control3 is used to hold a user password - encrypted by control1 
            '********************************************************************************************************

            'If Control2 = no, a dummy password will be stored in Control 3

            'Data will be encrypted by the GUI (Control1 - as unencrypted by default passphrase) + UserPassword (Control2 - as unencrypted by Control1)

            'Regard the password passed from the old database:
            'It was stored in the Keys to Send field of the Master Control record (double encrypted)
            'It started with either word "Yes" or "No " to indicate if a original password existed or not
            'If "No " some random numbers were be added following the "No "


            EncryptionClass.ResetDecryptionPassPhrase()
            EncryptionClass.UpdateDecryptionPassPhrase("LoadFromAccess" & My.Computer.Name.ToString)

            Dim YesOrNoPlusOriginalPasswordInPlainText As String = EncryptionClass.Decrypt(OriginalTable1Entries(0).Parameters)

            ' added in Push2Run - possible correction for A Form Filler
            If YesOrNoPlusOriginalPasswordInPlainText = String.Empty Then
                YesOrNoPlusOriginalPasswordInPlainText = "No " & GenerateRandomPassword(10)
            End If

            'If (YesOrNoPlusOriginalPasswordInPlainText.Length > 3) AndAlso (YesOrNoPlusOriginalPasswordInPlainText.StartsWith("Yes") OrElse YesOrNoPlusOriginalPasswordInPlainText.StartsWith("No ")) Then
            'Else
            '    Exit Try 'someone is screwing around
            'End If

            Dim Master_Password As String = GenerateRandomPassword(36)
            SQLcommand.Parameters.AddWithValue("@Control1", EncryptionClass.Encrypt(Master_Password)) 'encrypted with default pass phrase

            EncryptionClass.ResetEncryptionPassPhrase()
            EncryptionClass.UpdateEncryptionPassPhrase(Master_Password)  'Control2 and Control3 to be encrypted with the help of the Master_Password

            SQLcommand.Parameters.AddWithValue("@Control2", EncryptionClass.Encrypt(YesOrNoPlusOriginalPasswordInPlainText.Remove(3) & GenerateRandomPassword(7)))
            SQLcommand.Parameters.AddWithValue("@Control3", EncryptionClass.Encrypt(YesOrNoPlusOriginalPasswordInPlainText.Remove(0, 3)))

            EncryptionClass.UpdateEncryptionPassPhrase(YesOrNoPlusOriginalPasswordInPlainText.Remove(0, 3)) 'Data will be encrypted with Master_Password + Original Password

            SQLcommand.ExecuteNonQuery()

            SQLcommand.Dispose()
            SQLconnect.Close()

            'Step 2: some additional cleanup
            OriginalTable1Entries(0).Parameters = String.Empty

            LoadWasOK = True

        Catch ex As Exception

            LoadWasOK = False

        End Try

        Return LoadWasOK

    End Function

End Module
