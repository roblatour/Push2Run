
Imports System.IO

Module ModManageCards

    Friend XML_Path_Name As String
    Friend Const gPush2RunExtention As String = ".p2r"

    Friend Structure CardStructure
        Friend Description As String
        Friend ListenFor As String
        Friend Open As String
        Friend StartDirectory As String
        Friend Parameters As String
        Friend StartWithAdminPrivileges As Boolean
        Friend StartingWindowState As Integer
        Friend KeysToSend As String
    End Structure

    Friend Sub SetupToManageCards()

        XML_Path_Name = System.IO.Path.GetTempPath & "Push2Run"

        CreateFullDirectory(XML_Path_Name)

        DeleteAllTempFiles()

        InitalizeCard()

    End Sub

    Friend Sub DeleteAllTempFiles()

        Try
            Dim Folder As String = XML_Path_Name

            If IO.Directory.Exists(Folder) Then
                For Each _file As String In IO.Directory.GetFiles(Folder, "*" & gPush2RunExtention)
                    IO.File.Delete(_file)
                Next
                For Each _file As String In IO.Directory.GetFiles(Folder, "*" & gPush2RunExtention.ToUpper)
                    IO.File.Delete(_file)
                Next
            End If
        Catch ex As Exception

        End Try

    End Sub

    Friend Function InitalizeCard() As CardStructure

        Dim CurrentCard As CardStructure

        With CurrentCard

            .Description = String.Empty
            .ListenFor = String.Empty
            .Open = String.Empty
            .StartDirectory = String.Empty
            .Parameters = String.Empty
            .StartWithAdminPrivileges = False
            .StartingWindowState = 0
            .KeysToSend = String.Empty

        End With

        Return CurrentCard

    End Function

    Friend Sub CreateFullDirectory(ByVal DirectoryName As String)

        ' min length = 3 ; example "c:\"
        If DirectoryName.Length < 4 Then Exit Sub

        Dim x As Integer
        Dim ws As String = Microsoft.VisualBasic.Left(DirectoryName, 3)
        For x = 4 To Len(DirectoryName)
            If Mid(DirectoryName, x, 1) = "\" Then
                Try
                    If Not System.IO.Directory.Exists(ws) Then

                        ' v4.6
                        ' System.IO.Directory.CreateDirectory(ws)  

                        Dim workingDir As New DirectoryInfo(ws)
                        workingDir.Create()
                        workingDir.Refresh()

                    End If

                Catch ex As Exception
                End Try
            End If
            ws &= Mid(DirectoryName, x, 1)
        Next

        If DirectoryName.EndsWith("\") Then
        Else
            Try
                ' v4.6
                ' System.IO.Directory.CreateDirectory(ws)

                Dim workingDir As New DirectoryInfo(ws)
                workingDir.Create()
                workingDir.Refresh()

            Catch ex As Exception
            End Try
        End If

    End Sub

    Friend Function SaveCard(ByVal CardToSave As CardStructure) As String

        Dim ReturnValue As String = String.Empty

        'v2.0.2 'filter out invalid filename characters

        Dim XML_File_Name As String = CardToSave.Description.Trim & gPush2RunExtention
        XML_File_Name = String.Join("-", XML_File_Name.Split("\/" & IO.Path.GetInvalidFileNameChars()))  'filter out invalid filename characters
        Dim XML_Full_File_Name As String = XML_Path_Name & "\" & XML_File_Name.Replace(" ", "_") 'v4.3 added replace; saved filename will have underscores in place of spaces in the filenam

        Try

            Dim Card As CardClass = New CardClass

            With Card

                .Description = CardToSave.Description
                .ListenFor = CardToSave.ListenFor
                .Open = CardToSave.Open
                .StartDirectory = CardToSave.StartDirectory
                .Parameters = CardToSave.Parameters
                .StartWithAdminPrivileges = CardToSave.StartWithAdminPrivileges
                .StartingWindowState = CardToSave.StartingWindowState
                .KeysToSend = CardToSave.KeysToSend

            End With

            XML.ObjectXMLSerializer(Of CardClass).Save(Card, XML_Full_File_Name)

            Card = Nothing

            ReturnValue = XML_Full_File_Name

        Catch ex As Exception

            MsgBox(ex.Message.ToString)
            Call MsgBox("Unable to save the definition for the following card: '" & CardToSave.Description, MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly, "Push2Run - Warning")

        End Try

        Return ReturnValue

    End Function

    Friend Function LoadACard(ByVal XML_File_Name As String) As CardClass

        Dim LoadedCard As New CardClass

        With LoadedCard
            .Description = String.Empty
            .ListenFor = String.Empty
            .Open = String.Empty
            .StartDirectory = String.Empty
            .Parameters = String.Empty
            .StartWithAdminPrivileges = False
            .StartingWindowState = 0
            .KeysToSend = String.Empty
        End With

        Try

            If System.IO.File.Exists(XML_File_Name) Then

                '.p2r file format prior to version 3.2
                LoadedCard = XML.ObjectXMLSerializer(Of CardClass).Load(XML_File_Name)
                If LoadedCard Is Nothing Then
                    Call MsgBox("Unable to load definition from " & XML_File_Name, MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly, "Push2Run - Warning")
                End If

            Else

                Call MsgBox("Unable to find the file: " & XML_File_Name, MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly, "Push2Run - Warning")

            End If

        Catch ex As Exception

        End Try

        Return LoadedCard

    End Function

End Module
