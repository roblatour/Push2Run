Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports Microsoft.Win32
'Imports IWshRuntimeLibrary


Partial Public Class WindowAddChange

#Region "Add Help to Title Bar"

    ''ref http://stackoverflow.com/questions/1009983/help-button

    'Private Const WS_EX_CONTEXTHELP As UInteger = &H400
    'Private Const WS_MINIMIZEBOX As UInteger = &H20000
    'Private Const WS_MAXIMIZEBOX As UInteger = &H10000
    'Private Const GWL_STYLE As Integer = -16
    'Private Const GWL_EXSTYLE As Integer = -20
    'Private Const SWP_NOSIZE As Integer = &H1
    'Private Const SWP_NOMOVE As Integer = &H2fruna
    'Private Const SWP_NOZORDER As Integer = &H4
    'Private Const SWP_FRAMECHANGED As Integer = &H20
    'Private Const WM_SYSCOMMAND As Integer = &H112
    'Private Const SC_CONTEXTHELP As Integer = &HF180

    '<DllImport("user32.dll")> _
    'Private Shared Function GetWindowLong(ByVal hwnd As IntPtr, ByVal index As Integer) As UInteger
    'End Function

    '<DllImport("user32.dll")> _
    'Private Shared Function SetWindowLong(ByVal hwnd As IntPtr, ByVal index As Integer, ByVal newStyle As UInteger) As Integer
    'End Function

    '<DllImport("user32.dll")> _
    'Private Shared Function SetWindowPos(ByVal hwnd As IntPtr, ByVal hwndInsertAfter As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal width As Integer, ByVal height As Integer, _
    ' ByVal flags As UInteger) As Boolean
    'End Function

    'Protected Overrides Sub OnSourceInitialized(ByVal e As EventArgs)
    '    MyBase.OnSourceInitialized(e)
    '    Dim hwnd As IntPtr = New System.Windows.Interop.WindowInteropHelper(Me).Handle
    '    Dim styles As UInteger = GetWindowLong(hwnd, GWL_STYLE)

    '    'styles = styles And &HFFFFFFFFUI Xor (WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
    '    styles = styles And &HFFFFFFFFUI Xor WS_MAXIMIZEBOX

    '    SetWindowLong(hwnd, GWL_STYLE, styles)
    '    styles = GetWindowLong(hwnd, GWL_EXSTYLE)
    '    styles = styles Or WS_EX_CONTEXTHELP
    '    SetWindowLong(hwnd, GWL_EXSTYLE, styles)
    '    SetWindowPos(hwnd, IntPtr.Zero, 0, 0, 0, 0, _
    '    SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED)
    '    DirectCast(PresentationSource.FromVisual(Me), HwndSource).AddHook(AddressOf HelpHook)
    'End Sub

    'Private Function HelpHook(ByVal hwnd As IntPtr, ByVal msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr, ByRef handled As Boolean) As IntPtr
    '    If msg = WM_SYSCOMMAND AndAlso (CInt(wParam) And &HFFF0) = SC_CONTEXTHELP Then

    '        OpenAWebPage(gWebPageHelpChangeWindow)

    '        handled = True

    '    End If
    '    Return IntPtr.Zero
    'End Function

#End Region

    Public Const GW_HWNDNEXT As Integer = 2
    Public Const GW_CHILD As Integer = 5

    Private NoTitleWindow As String
    Private Enum StatusValues
        SwitchOff = 0
        SwitchOn = 1
        NoSwitch = 2
    End Enum

    Private HoldDescription As String = String.Empty
    Private HoldListenFor As String = String.Empty
    Private HoldOpen As String = String.Empty
    Private HoldParameters As String = String.Empty
    Private HoldStartIn As String = String.Empty
    Private HoldAdmin As Boolean = False
    Private HoldWindowState As Integer = 0
    Private HoldKeysToSend As String = String.Empty

    Private FormIsLoading As Boolean = True



    <Obfuscation(Feature:="virtualization", Exclude:=False)>
    Private Sub frmAddChange_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Loaded

        SeCursor(CursorState.Wait)

        Me.Height = LastSizeOfAddChangeWindow.Height
        Me.Width = LastSizeOfAddChangeWindow.Width

        'if the Add Change Window has not been opened before, open it near the main window, otherwise open it where it was last opened
        If LastLocationOfAddChangeWindow = New Point(-1, -1) Then
            Me.Top = My.Settings.Top + 100
            Me.Left = My.Settings.Left + 100
        Else
            Me.Top = LastLocationOfAddChangeWindow.X
            Me.Left = LastLocationOfAddChangeWindow.Y
        End If

        Me.ShowInTaskbar = False

        MakeTopMost(SafeNativeMethods.FindWindow(Nothing, Me.Title), My.Settings.AlwaysOnTop)

        With gCurrentlySelectedRow

            HoldDescription = .Description
            HoldListenFor = .ListenFor
            HoldOpen = .Open
            HoldParameters = .Parameters
            HoldStartIn = .StartIn
            HoldAdmin = .Admin
            HoldWindowState = .StartingWindowState
            HoldKeysToSend = .KeysToSend

            Select Case gLoadingAddChange

                Case Is = "Add"
                    tbDescription.Text = String.Empty
                    tbListenFor.Text = String.Empty
                    tbOpen.Text = String.Empty
                    tbParameters.Text = String.Empty
                    tbStartin.Text = String.Empty
                    cbAdmin.IsChecked = False
                    cbWindowState.Text = "Normal"
                    tbKeysToSend.Text = String.Empty

                Case Is = "Edit"
                    tbDescription.Text = .Description
                    tbListenFor.Text = .ListenFor
                    tbOpen.Text = .Open
                    tbParameters.Text = .Parameters
                    tbStartin.Text = .StartIn
                    cbAdmin.IsChecked = .Admin
                    cbWindowState.Text = .StartingWindowState.ConvertStartingWindowStateToAString
                    tbKeysToSend.Text = .KeysToSend

            End Select

        End With

        NoTitleWindow = Chr(149)

        tbDescription.Focus()
        tbDescription.CaretIndex = tbDescription.Text.Length

        gReturnFromAddChange = "Cancel"

        Me.Activate()

        FormIsLoading = False

        SeCursor(CursorState.Normal)

    End Sub

    '    uri = New Uri("/Push2Run;component/Resources/switchoff.png", UriKind.Relative)
    '    Me.SwitchImage.Source = New System.Windows.Media.Imaging.BitmapImage(uri)

    '   Dim uri = New Uri("/Push2Run;component/Resources/greendot.png", UriKind.Relative)
    '   Greendot.Source = New System.Windows.Media.Imaging.BitmapImage(uri)

    <Obfuscation(Feature:="virtualization", Exclude:=False)>
    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click

        Dim CleanedUpDescription As String = CleanUpWhiteAndDuplicatedSpaces(tbDescription.Text)

        If CleanedUpDescription.ToUpper = gMasterSwitch.ToUpper Then

            Beep()
            Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "You cannot add an entry with a description of 'Master Switch'.", "Push2Run - Info", MessageBoxButton.OK, MessageBoxImage.Warning)

            Exit Sub

        End If

        Dim AllPhrasesToListenFor() As String = tbListenFor.Text.Split(vbCr)
        For Each phrase As String In AllPhrasesToListenFor

            phrase = phrase.Replace(" ", "").Trim

            If phrase = "*$" Then
                Beep()
                Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "A 'Listen for' phrase may not have a '*' and a '$' without words, numbers, and/or symbols between them.", "Push2Run - Info", MessageBoxButton.OK, MessageBoxImage.Warning)
                Exit Sub
            End If

            Dim i1 As Integer = phrase.IndexOf("*")
            Dim i2 As Integer = phrase.IndexOf("$")
            If (i1 >= 0) And (i2 >= 0) Then
                If i1 > i2 Then
                    Beep()
                    Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "A 'Listen for' phrase may not have a '*' and a '$' where the '$' comes before the '*'.", "Push2Run - Info", MessageBoxButton.OK, MessageBoxImage.Warning)
                    Exit Sub
                End If
            End If

        Next

        If tbKeysToSend.Text.Length > 0 Then

            If System.String.IsNullOrWhiteSpace(tbKeysToSend.Text) Then

                If TopMostMessageBox(gCurrentOwner, "Warning: the 'Keys to send' field contains all white spaces." & vbCrLf & vbCrLf &
                                   "Please click 'Yes' if this is what is intended, or 'No' to remove the white spaces.", "Push2Run - Info", MessageBoxButton.YesNo, MessageBoxImage.Warning, vbYes) = vbNo Then
                    tbKeysToSend.Text = String.Empty
                End If

            End If

        End If


        If tbKeysToSend.Text.Length > 0 Then

            If (cbWindowState.Text = "Hidden") OrElse (cbWindowState.Text = "Minimized") Then
                Beep()
                Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "Keys cannot be sent to a minimized or hidden window.", "Push2Run - Info", MessageBoxButton.OK, MessageBoxImage.Warning)
                Exit Sub
            End If

        End If

        If (tbKeysToSend.Text.Length > 0) AndAlso (cbAdmin.IsChecked) Then
            Beep()
            Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "Note: Push2Run cannot send keys to a program running with admin privileges unless Push2Run is itself running with admin privileges.", "Push2Run - Info", MessageBoxButton.OK, MessageBoxImage.Information)
            ' this  is a warning only, so an Exit Sub is not used here
        End If

        Try

            With gCurrentlySelectedRow

                .Description = CleanedUpDescription
                .ListenFor = StringSort(tbListenFor.Text.Replace(vbLf, vbCrLf))
                .Open = tbOpen.Text
                .Parameters = tbParameters.Text
                .StartIn = tbStartin.Text
                .Admin = cbAdmin.IsChecked
                .StartingWindowState = cbWindowState.Text.ConvertStartingWindowStateToANumber
                .KeysToSend = tbKeysToSend.Text

                If (.Description.Trim.Length = 0) AndAlso (.ListenFor.Trim.Length = 0) AndAlso (.Open.Trim.Length = 0) AndAlso (.Parameters.Trim.Length = 0) AndAlso (.StartIn.Length = 0) AndAlso (Not .Admin) AndAlso (.KeysToSend.Length = 0) Then

                    Dim Message As String = String.Empty

                    If (tbDescription.Text.Length > 0) OrElse (tbListenFor.Text.Length > 0) OrElse (tbOpen.Text.Length > 0) OrElse (tbParameters.Text.Length > 0) OrElse (tbStartin.Text.Length > 0) OrElse (tbKeysToSend.Text.Length > 0) Then
                        Message = "Effectively no information has been entered, would you like to cancel?"
                    Else
                        Message = "No information has been entered, would you like to cancel?"
                    End If

                    If TopMostMessageBox(gCurrentOwner, Message, "Push2Run - Info", MessageBoxButton.YesNo, MessageBoxImage.Question) = MessageBoxResult.Yes Then
                        CancelLogic()
                    End If

                    Exit Sub

                End If


                ' I don't have confidence in that these edits will not raise false positives
                ' levaing them out for now
                'Try

                '    'if regex groups are used in the listen field, validate regex  

                '    Dim RegexRule = New Regex(gCurrentlySelectedRow.ListenFor, RegexOptions.IgnoreCase)

                '    If (RegexRule.GetGroupNames.Count > 1) Then

                '        If Not ValidateRegex("Open", .Open) Then Exit Sub
                '        If Not ValidateRegex("Parameters", .Parameters) Then Exit Sub
                '        If Not ValidateRegex("Keys to send", .KeysToSend) Then Exit Sub

                '    End If

                'Catch ex As Exception
                'End Try


                'Validate .open and .startin

                ' look for the Push2Run .start directory

                .StartIn = .StartIn.Trim

                If .StartIn.Trim.Length > 0 Then
                    If Directory.Exists(Environment.ExpandEnvironmentVariables(.StartIn.Trim)) Then  'v3.4.2
                        ' continue on 
                    Else
                        Beep()
                        Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "Start directory does not exist.", "Push2Run", MessageBoxButton.OK, MessageBoxImage.Warning)
                        Exit Sub
                    End If
                End If

                Try

                    'Validate .Open field

                    If .Open.ToUpper.Trim = "DESKTOP" Then

                        .Open = "Desktop"

                        If cbAdmin.IsChecked Then
                            Beep()
                            Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "The desktop cannot be opened with admin privileges", "Push2Run - Info", MessageBoxButton.OK, MessageBoxImage.Warning)
                            Exit Sub
                        End If

                        Exit Try

                    End If

                    If (.Open.ToUpper.Trim = "ACTIVEWINDOW") Then
                        .Open = "Active Window"
                    End If

                    If (.Open.ToUpper.Trim = "ACTIVE WINDOW") Then

                        If cbAdmin.IsChecked Then
                            Beep()
                            Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "The active window cannot be opened with admin privileges", "Push2Run - Info", MessageBoxButton.OK, MessageBoxImage.Warning)
                            Exit Sub
                        End If

                        Exit Try

                    End If

                    If .Open.ToUpper.Trim = "MQTT" Then

                        .ListenFor = .ListenFor.Trim
                        .Parameters = .Parameters.Trim

                        Dim Topic As String = String.Empty
                        Dim PayLoad As String = String.Empty

                        Dim Result As String = ValidatePublish(.Parameters, Topic, PayLoad)

                        If Result = "ok" Then

                        ElseIf Result.StartsWith("Warning:") Then
                            Beep()
                            Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner, Result, "Push2Run - Warning", MessageBoxButton.OK, MessageBoxImage.Warning)

                        ElseIf Result.StartsWith("Warnings:") Then
                            Beep()
                            Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner, Result, "Push2Run - Warnings", MessageBoxButton.OK, MessageBoxImage.Warning)

                        Else
                            Beep()
                            Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner, Result, "Push2Run - Info", MessageBoxButton.OK, MessageBoxImage.Information)
                            Exit Sub

                        End If

                        'For this next edit it is important to note MQTT topics are case sensitive so Weather/status <> weather/status
                        Dim PayloadFromListenForField As String = .ListenFor
                        PayloadFromListenForField = PayloadFromListenForField.Replace(Topic, "").Trim()

                        If .ListenFor.StartsWith(Topic) AndAlso (PayloadFromListenForField = PayLoad) Then
                            Beep()
                            Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "When using MQTT the 'Listen for' and 'Paramters' fields may not be the same.", "Push2Run - Info", MessageBoxButton.OK, MessageBoxImage.Warning)
                            Exit Sub
                        End If

                        Exit Try

                    End If

                    If .Open.ToLower.StartsWith("http") OrElse .Open.ToUpper.StartsWith("www") Then
                        Exit Try
                    End If

                    If .Open.ToLower.Contains("[everything") Then
                        Exit Try
                    End If

                    Dim ExecutableFileExtenstions() As String = {"", ".exe", ".bat", ".vbs", ".ps1"}

                    'search for executable program as entered or with one of the above extentions
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


                    'search for executable program in the system path as entered or with one of the above extentions
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

                        If File.Exists(AlternateSystem32FolderName & "\" & Filename & Filetype) Then

                            If TopMostMessageBox(gCurrentOwner, "It's complicated, but the program you likely want to open appears within Windows\System32 when viewed by Windows File Exporer - but it is really not there at all." & vbCrLf & vbCrLf &
                                               "Push2Run can correct for this by setting:" & vbCrLf & vbCrLf &
                                               "the 'Open' field in your Push2Run card to: " & vbCrLf & Filename & Filetype & vbCrLf & vbCrLf &
                                               "and" & vbCrLf & vbCrLf &
                                               "the 'Start Directory' field in your Push2Run card to: " & vbCrLf & AlternateSystem32FolderName & vbCrLf & vbCrLf &
                                               "Would you like Push2Run to do this now?",
                                               "Push2Run",
                                       MessageBoxButton.YesNo, MessageBoxImage.Question) = MessageBoxResult.Yes Then

                                .Open = Filename & Filetype
                                .StartIn = AlternateSystem32FolderName

                                'all is good
                                Exit Try

                            End If

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

                            If TopMostMessageBox(gCurrentOwner, "Push2Run found the program '" & TestProgramName & "' in the directory:" & vbCrLf & vbCrLf &
                                               .StartIn & vbCrLf & vbCrLf &
                                               "Do you want your Push2Run card to specify this directory?",
                                               "Push2Run",
                                       MessageBoxButton.YesNo, MessageBoxImage.Question) = MessageBoxResult.Yes Then

                                .Open = .StartIn & "\" & TestProgramName
                                .Open = .Open.Replace("\\", "\")

                            Else

                                .StartIn = String.Empty

                            End If

                            'its all good  (as the entry was found in the registry, the card will work with or without changes)
                            Exit Try

                        End If

                    End If


                    If TopMostMessageBox(gCurrentOwner, "The program to open couldn't be found." & vbCrLf & vbCrLf &
                                       "Would you like to save this infomation anyway?",
                                       "Push2Run",
                                       MessageBoxButton.YesNo, MessageBoxImage.Question) = MessageBoxResult.No Then Exit Sub


                    '' search for program in the common folders
                    'Dim Message As String = "The program to open is not found in the identifed folder." & vbCrLf & vbCrLf &
                    '                        "Would you like Push2Run to search for program in other common program folders on your system?" & vbCrLf & vbCrLf &
                    '                        "PLEASE NOTE: this could take up to a minute to run, please be patient"

                    'If TopMostMessageBox(gCurrentOwner,Message, "Push2Run", MessageBoxButton.YesNo, MessageBoxImage.Question) = MessageBoxResult.Yes Then

                    '    Dim results As List(Of String) = InventoryAllPrograms(Filename)
                    '    If results.Count = 1 Then
                    '        .Open = results(0)
                    '        Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner,"Card information had been automatically changed, please review and if fine click 'ok'")
                    '        Exit Sub
                    '    End If

                    '  End If

                Catch ex As Exception

                End Try

                gReturnFromAddChange = "OK"

                gReturnFromAddChangeDataChanged = (.Description <> HoldDescription) OrElse (.ListenFor <> HoldListenFor) OrElse (.Open <> HoldOpen) OrElse (.Parameters <> HoldParameters) OrElse (.StartIn <> HoldStartIn) OrElse
                                                  (.Admin <> HoldAdmin) OrElse (.StartingWindowState <> HoldWindowState) OrElse (.KeysToSend <> HoldKeysToSend)

            End With

            Me.Close()

        Catch ex As Exception

        End Try

    End Sub

    Private Function ValidateRegex(ByVal description As String, ByVal input As String) As Boolean

        Dim ReturnValue As Boolean = True

        Try

            ' the following will throw an exception if the input is not a valid regex expression
            Dim dummy = New Regex(input)

        Catch ex As Exception

            Beep()
            Dim dummy As MessageBoxResult = TopMostMessageBox(gCurrentOwner, "'" & description & "' contains an invalid Regex expresssion.", "Push2Run", MessageBoxButton.OK, MessageBoxImage.Warning)
            ReturnValue = False

        End Try

        Return ReturnValue

    End Function


    <Obfuscation(Feature:="virtualization", Exclude:=False)>
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click

        CancelLogic()

    End Sub

    <Obfuscation(Feature:="virtualization", Exclude:=False)>
    Private Sub CancelLogic()

        gReturnFromAddChange = "Cancel"
        gReturnFromAddChangeDataChanged = False
        Me.Close()

    End Sub

    <Obfuscation(Feature:="virtualization", Exclude:=False)>
    Private Sub btnHelp_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnHelp.Click

        Me.Dispatcher.Invoke(New OpenAWebPageDelegate(AddressOf OpenAWebPage), New Object() {gWebPageHelpChangeWindow})

    End Sub


    Private Sub WindowAddChange_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

        LastLocationOfAddChangeWindow = New Point(Me.Top, Me.Left)
        LastSizeOfAddChangeWindow = New Point(Me.Width, Me.Height)

    End Sub

    Private Sub ThisWindow_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged

        KeepHelpOnTop()

    End Sub

    Private Sub tb_PreviewDrop(sender As Object, e As DragEventArgs)

        With gCurrentlySelectedRow

            .Description = tbDescription.Text
            .ListenFor = tbListenFor.Text
            .Open = tbOpen.Text
            .Parameters = tbParameters.Text
            .StartIn = tbStartin.Text
            .Admin = cbAdmin.IsChecked
            .StartingWindowState = cbWindowState.Text.ConvertStartingWindowStateToANumber
            .KeysToSend = tbKeysToSend.Text

            Dim ShowPreview As Boolean = DropIntogCurrentlySelectedRow(e)

            If ShowPreview Then

                tbDescription.Text = .Description
                tbListenFor.Text = .ListenFor
                tbOpen.Text = .Open
                tbStartin.Text = .StartIn
                tbParameters.Text = .Parameters
                cbAdmin.IsChecked = .Admin
                cbWindowState.Text = .StartingWindowState.ConvertStartingWindowStateToAString
                tbKeysToSend.Text = .KeysToSend

            End If

        End With

    End Sub

    Private Sub tb_PreviewDragEnter(sender As Object, e As DragEventArgs)

        e.Effects = DragDropEffects.Copy
        e.Handled = True

    End Sub

    Private Sub tb_PreviewDragLeave(sender As Object, e As DragEventArgs)

        e.Effects = DragDropEffects.None
        e.Handled = True

    End Sub

    Private Sub WindowAddChange_StateChanged(sender As Object, e As EventArgs) Handles Me.StateChanged

        If sender.windowstate = WindowState.Minimized Then
            ' Beep()
            Me.WindowState = WindowState.Normal
        End If

    End Sub

End Class
