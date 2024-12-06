Imports System.ComponentModel
Imports System.Text

Public Class WindowSessionLog

    Private ScrollViewer As ScrollViewer

    Private Sub WindowLog_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded

        SessionLogIsOpen = True

        InitializeComponent()

        If LastLocationofSessionLog = Nothing Then
            LastLocationofSessionLog = New Point(Me.Top, Me.Left)
        Else
            Me.Top = LastLocationofSessionLog.X
            Me.Left = LastLocationofSessionLog.Y
        End If

        If LastSizeofSessionLog = Nothing Then
            LastSizeofSessionLog = New Size(Me.Width, Me.Height)
        Else
            Me.Width = LastSizeofSessionLog.Width
            Me.Height = LastSizeofSessionLog.Height
        End If

        lbSessionLog.ItemsSource = SessionLogListViewData

        AddHandler EventToUpdateTheLogWindowsSessionLog, New EventHandler(AddressOf UpdateTheLog)

        Dim border As Border = DirectCast(VisualTreeHelper.GetChild(lbSessionLog, 0), Border)
        ScrollViewer = DirectCast(VisualTreeHelper.GetChild(border, 0), ScrollViewer)

        SetAutoScrollStatus(My.Settings.AutoScroll)

        'Reload the window with the same top and selected items

        If IndexToBeSelectedOnReload >= 0 Then

            DragAndDropUnderway = False

            If ItemContentToBeScrolledToOnReload > String.Empty Then

                For Each workingitem In lbSessionLog.Items

                    If workingitem = ItemContentToBeScrolledToOnReload Then

                        lbSessionLog.SelectedItem = workingitem
                        lbSessionLog.ScrollIntoView(lbSessionLog.SelectedItem)
                        Exit For

                    End If

                Next

                'scroll into view the top most row from the prior session
                ScrollViewer.ScrollToBottom()
                lbSessionLog.ScrollIntoView(lbSessionLog.SelectedItem)

                'set as the selected row the selected row of the prior session
                lbSessionLog.SelectedIndex = IndexToBeSelectedOnReload
                lbSessionLog.SelectedIndex = -1  ' selection is removed so drag and drop can be redone on the same row

            End If

        Else

            UpdateTheLog()

        End If

        SeCursor(CursorState.Normal)

    End Sub

    Private Sub UpdateTheLog()

        ' Updates the log to show the most current item at the bottom

        If My.Settings.AutoScroll Then
            ScrollViewer.ScrollToBottom()
        End If

    End Sub

    Private Sub WindowSessionLog_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

        RaiseAnEventToCloseSessionLog()

        RemoveHandler EventToUpdateTheLogWindowsSessionLog, AddressOf UpdateTheLog

        LastLocationofSessionLog = New Point(Me.Top, Me.Left)
        LastSizeofSessionLog = New Size(Me.Width, Me.Height)

        SessionLogIsOpen = False

    End Sub

    Private Sub WindowSessionLog_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged

        KeepHelpOnTop()

    End Sub

#Region "Drag and Drop"

    ' ref https://www.youtube.com/watch?v=ZIbOY9Z9fKI
    ' ref https://stackoverflow.com/questions/14731304/why-doesnt-doubleclick-event-fire-after-mousedown-event-on-same-element-fires

    ' allows the user to drag a <Body> entry out of the session log and drop it into the boss window 

    Private Sub lbSessionLog_PreviewMouseDown(sender As Object, e As MouseButtonEventArgs) Handles lbSessionLog.PreviewMouseDown

        DragAndDropUnderway = True

    End Sub

    Private Sub lbSessionLog_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles lbSessionLog.SelectionChanged

        If DragAndDropUnderway Then

            DragAndDropUnderway = False


            Try

                Dim item As ListBox = sender

                If item IsNot Nothing Then

                    Dim ItemSelectedText As String = lbSessionLog.SelectedItem.ToString.Trim

                    If ItemSelectedText > String.Empty Then

                        If ItemSelectedText.Contains("- <Body>") Then

                            ' Find the top most displayed textbox 
                            ' ref https://stackoverflow.com/questions/2926722/get-first-visible-item-in-wpf-listview-c-sharp

                            'Dim hitTest As HitTestResult = VisualTreeHelper.HitTest(lbSessionLog, New Point(15, 15))
                            'Dim TopOfTheListItem As System.Windows.Controls.ListBoxItem = TryCast(GetListBoxItemFromEvent(Nothing, hitTest.VisualHit), System.Windows.Controls.ListBoxItem)

                            Dim hitTest As HitTestResult
                            Dim TopOfTheListItem As System.Windows.Controls.ListBoxItem

                            Dim XCoordinate As Integer = 15
                            Dim YCordinate As Integer = 15

                            Try

                                Do

                                    hitTest = VisualTreeHelper.HitTest(lbSessionLog, New Point(XCoordinate, YCordinate))
                                    TopOfTheListItem = TryCast(GetListBoxItemFromEvent(Nothing, hitTest.VisualHit), System.Windows.Controls.ListBoxItem)
                                    YCordinate += 5

                                Loop While TopOfTheListItem.Content.trim = String.Empty

                                ItemContentToBeScrolledToOnReload = TopOfTheListItem.Content

                            Catch ex As Exception

                                ItemContentToBeScrolledToOnReload = String.Empty

                            End Try

                            IndexToBeSelectedOnReload = lbSessionLog.SelectedIndex

                            Dim TempItem As New System.Windows.Forms.Label
                            TempItem.Text = ItemSelectedText.Remove(0, ItemSelectedText.IndexOf("- <Body>") + "- <Body>".Length).Trim
                            TempItem.DoDragDrop(TempItem.Text, DragDropEffects.Copy)

                            ' after the drag drop item is processed in the boss 
                            ' this form will be shutdown by the Boss 
                            ' this allows the release of the dragdrop item

                            TempItem.Dispose()

                            Exit Sub

                        End If

                    End If

                End If

            Catch ex As Exception

            End Try


        End If

        SetAutoScrollStatus(My.Settings.AutoScroll)

        DragAndDropUnderway = False

    End Sub

    Private Function GetListBoxItemFromEvent(sender As Object, originalSource As Object) As System.Windows.Controls.ListBoxItem

        Dim depObj As DependencyObject = TryCast(originalSource, DependencyObject)

        If depObj IsNot Nothing Then

            ' go up the visual hierarchy until we find the list view item the click came from  
            ' the click might have been on the grid or column headers so we need to cater for this  

            Dim current As DependencyObject = depObj

            While current IsNot Nothing ' AndAlso current <> 
                Dim ListboxItem As System.Windows.Controls.ListBoxItem = TryCast(current, System.Windows.Controls.ListBoxItem)
                If ListboxItem IsNot Nothing Then
                    Return ListboxItem
                End If
                current = VisualTreeHelper.GetParent(current)
            End While

        End If

        Return Nothing
    End Function

    Private Sub lbSessionLog_PreviewMouseUp(sender As Object, e As MouseButtonEventArgs) Handles lbSessionLog.PreviewMouseUp

        DragAndDropUnderway = False

    End Sub


    Private Sub AutoScroll_PreviewMouseDown(sender As Object, e As MouseButtonEventArgs) Handles imgAutoScroll.PreviewMouseDown

        SetAutoScrollStatus(Not My.Settings.AutoScroll)

    End Sub

    Private Sub SetAutoScrollStatus(ByVal Status As Boolean)

        My.Settings.AutoScroll = Status

        Dim c As New ImageSourceConverter()

        If My.Settings.AutoScroll Then
            imgAutoScroll.Source = CType(c.ConvertFrom(New Uri("pack://application:,,,/Resources/switchon.png")), ImageSource)
            ScrollViewer.ScrollToBottom()
        Else
            imgAutoScroll.Source = CType(c.ConvertFrom(New Uri("pack://application:,,,/Resources/switchoff.png")), ImageSource)
        End If

    End Sub

    Private Sub CopyToClipboard_Click(sender As Object, e As RoutedEventArgs)

        Dim sb As New StringBuilder

        For Each IndividualLine As String In lbSessionLog.Items

            sb.Append(IndividualLine & vbCrLf)

        Next

        My.Computer.Clipboard.SetText(sb.ToString)

        sb = Nothing

    End Sub

    Private Sub ClearSessionLog_Click(sender As Object, e As RoutedEventArgs)


        If TopMostMessageBox(Me, "Are you sure you want to clear the Session log?",
                                   "Push2Run - Clear the Session log", MessageBoxButton.YesNo, MessageBoxImage.Question, System.Windows.MessageBoxOptions.None) = MessageBoxResult.Yes Then

            SessionLogListViewData.Clear()

        End If


    End Sub

#End Region

End Class

