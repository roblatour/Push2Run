Module modEverything

    ' ref: 


    Public Declare Unicode Function Everything_32Bit_x86_SetSearchW Lib "C:\Program Files (x86)\Push2Run\Everything32.dll" Alias "Everything_SetSearchW" (ByVal search As String) As UInt32
    Public Declare Unicode Function Everything_32Bit_x86_SetRequestFlags Lib "C:\Program Files (x86)\Push2Run\Everything32.dll" Alias "Everything_SetRequestFlags" (ByVal dwRequestFlags As UInt32) As UInt32
    Public Declare Unicode Function Everything_32Bit_x86_QueryW Lib "C:\Program Files (x86)\Push2Run\Everything32.dll" Alias "Everything_QueryW" (ByVal bWait As Integer) As Integer
    Public Declare Unicode Function Everything_32Bit_x86_GetNumResults Lib "C:\Program Files (x86)\Push2Run\Everything32.dll" Alias "Everything_GetNumResults" () As UInt32
    'Public Declare Unicode Function Everything_32Bit_x86_GetResultFileNameW Lib "C:\Program Files (x86)\Push2Run\Everything32.dll" Alias "Everything_GetResultFileNameW" (ByVal index As UInt32) As IntPtr
    'Public Declare Unicode Function Everything_32Bit_x86_GetLastError Lib "C:\Program Files\Push2Run (x86)\Everything32.dll" Alias "Everything_GetLastError" () As UInt32
    Public Declare Unicode Function Everything_32Bit_x86_GetResultFullPathNameW Lib "C:\Program Files (x86)\Push2Run\Everything32.dll" Alias "Everything_GetResultFullPathNameW" (ByVal index As UInt32, ByVal buf As System.Text.StringBuilder, ByVal size As UInt32) As UInt32
    'Public Declare Unicode Function Everything_32Bit_x86_GetResultSize Lib "C:\Program Files\Push2Run (x86)\Everything32.dll" Alias "Everything_GetResultSize" (ByVal index As UInt32, ByRef size As UInt64) As Integer
    'Public Declare Unicode Function Everything_32Bit_x86_GetResultDateModified Lib "C:\Program Files (x86)\Push2Run\Everything32.dll" Alias "Everything_GetResultDateModified" (ByVal index As UInt32, ByRef ft As UInt64) As Integer


    Public Declare Unicode Function Everything_32Bit_SetSearchW Lib "C:\Program Files\Push2Run\Everything32.dll" Alias "Everything_SetSearchW" (ByVal search As String) As UInt32
    Public Declare Unicode Function Everything_32Bit_SetRequestFlags Lib "C:\Program Files\Push2Run\Everything32.dll" Alias "Everything_SetRequestFlags" (ByVal dwRequestFlags As UInt32) As UInt32
    Public Declare Unicode Function Everything_32Bit_QueryW Lib "C:\Program Files\Push2Run\Everything32.dll" Alias "Everything_QueryW" (ByVal bWait As Integer) As Integer
    Public Declare Unicode Function Everything_32Bit_GetNumResults Lib "C:\Program Files\Push2Run\Everything32.dll" Alias "Everything_GetNumResults" () As UInt32
    'Public Declare Unicode Function Everything_32Bit_GetResultFileNameW Lib "C:\Program Files\Push2Run\Everything32.dll" Alias "Everything_GetResultFileNameW" (ByVal index As UInt32) As IntPtr
    'Public Declare Unicode Function Everything_32Bit_GetLastError Lib "C:\Program Files\Push2Run\Everything32.dll" Alias "Everything_GetLastError" () As UInt32
    Public Declare Unicode Function Everything_32Bit_GetResultFullPathNameW Lib "C:\Program Files\Push2Run\Everything32.dll" Alias "Everything_GetResultFullPathNameW" (ByVal index As UInt32, ByVal buf As System.Text.StringBuilder, ByVal size As UInt32) As UInt32
    'Public Declare Unicode Function Everything_32Bit_GetResultSize Lib "C:\Program Files\Push2Run\Everything32.dll" Alias "Everything_GetResultSize" (ByVal index As UInt32, ByRef size As UInt64) As Integer
    'Public Declare Unicode Function Everything_32Bit_GetResultDateModified Lib "C:\Program Files\Push2Run\Everything32.dll" Alias "Everything_GetResultDateModified" (ByVal index As UInt32, ByRef ft As UInt64) As Integer

    Public Declare Unicode Function Everything_64Bit_SetSearchW Lib "C:\Program Files\Push2Run\Everything64.dll" Alias "Everything_SetSearchW" (ByVal search As String) As UInt32
    Public Declare Unicode Function Everything_64Bit_SetRequestFlags Lib "C:\Program Files\Push2Run\Everything64.dll" Alias "Everything_SetRequestFlags" (ByVal dwRequestFlags As UInt32) As UInt32
    Public Declare Unicode Function Everything_64Bit_QueryW Lib "C:\Program Files\Push2Run\Everything64.dll" Alias "Everything_QueryW" (ByVal bWait As Integer) As Integer
    Public Declare Unicode Function Everything_64Bit_GetNumResults Lib "C:\Program Files\Push2Run\Everything64.dll" Alias "Everything_GetNumResults" () As UInt32
    'Public Declare Unicode Function Everything_64Bit_GetResultFileNameW Lib "C:\Program Files\Push2Run\Everything64.dll" Alias "Everything_GetResultFileNameW" (ByVal index As UInt32) As IntPtr
    'Public Declare Unicode Function Everything_64Bit_GetLastError Lib "C:\Program Files\Push2Run\Everything64.dll" Alias "Everything_GetLastError" () As UInt32
    Public Declare Unicode Function Everything_64Bit_GetResultFullPathNameW Lib "C:\Program Files\Push2Run\Everything64.dll" Alias "Everything_GetResultFullPathNameW" (ByVal index As UInt32, ByVal buf As System.Text.StringBuilder, ByVal size As UInt32) As UInt32
    'Public Declare Unicode Function Everything_64Bit_GetResultSize Lib "C:\Program Files\Push2Run\Everything64.dll" Alias "Everything_GetResultSize" (ByVal index As UInt32, ByRef size As UInt64) As Integer
    'Public Declare Unicode Function Everything_64Bit_GetResultDateModified Lib "C:\Program Files\Push2Run\Everything64.dll" Alias "Everything_GetResultDateModified" (ByVal index As UInt32, ByRef ft As UInt64) As Integer

    Public Const EVERYTHING_REQUEST_FILE_NAME = &H1
    Public Const EVERYTHING_REQUEST_PATH = &H2
    'Public Const EVERYTHING_REQUEST_FULL_PATH_AND_FILE_NAME = &H4
    'Public Const EVERYTHING_REQUEST_EXTENSION = &H8
    'Public Const EVERYTHING_REQUEST_SIZE = &H10
    'Public Const EVERYTHING_REQUEST_DATE_CREATED = &H20
    'Public Const EVERYTHING_REQUEST_DATE_MODIFIED = &H40
    'Public Const EVERYTHING_REQUEST_DATE_ACCESSED = &H80
    'Public Const EVERYTHING_REQUEST_ATTRIBUTES = &H100
    'Public Const EVERYTHING_REQUEST_FILE_LIST_FILE_NAME = &H200
    'Public Const EVERYTHING_REQUEST_RUN_COUNT = &H400
    'Public Const EVERYTHING_REQUEST_DATE_RUN = &H800
    'Public Const EVERYTHING_REQUEST_DATE_RECENTLY_CHANGED = &H1000
    'Public Const EVERYTHING_REQUEST_HIGHLIGHTED_FILE_NAME = &H2000
    'Public Const EVERYTHING_REQUEST_HIGHLIGHTED_PATH = &H4000
    'Public Const EVERYTHING_REQUEST_HIGHLIGHTED_FULL_PATH_AND_FILE_NAME = &H8000

    'Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

    '    Everything_SetSearchW(TextBox1.Text)
    '    'Everything_SetRequestFlags(EVERYTHING_REQUEST_FILE_NAME Or EVERYTHING_REQUEST_PATH Or EVERYTHING_REQUEST_SIZE Or EVERYTHING_REQUEST_DATE_MODIFIED)
    '    Everything_SetRequestFlags(EVERYTHING_REQUEST_FILE_NAME)
    '    Everything_QueryW(1)

    '    Dim NumResults As UInt32
    '    Dim i As UInt32
    '    Dim filename As New System.Text.StringBuilder(260)
    '    Dim size As UInt64
    '    Dim ft dm As UInt64
    '    Dim DateModified As System.DateTime

    '    NumResults = Everything_GetNumResults()

    '    ListBox1.Items.Clear()

    '    If NumResults > 0 Then
    '        For i = 0 To NumResults - 1

    '            Everything_GetResultFullPathNameW(i, filename, filename.Capacity)
    '            Everything_GetResultSize(i, size)
    '            Everything_GetResultDateModified(i, ft dm)

    '            ' Everything uses &HFFFFFFFFFFFFFFFFUL for unknown dates
    '            ' System.DateTime.FromFileTime does not like this value
    '            ' so set the DateModified to Nothing when Everything returns &HFFFFFFFFFFFFFFFFUL
    '            If ftdm = &HFFFFFFFFFFFFFFFFUL Then
    '                DateModified = Nothing
    '            Else
    '                DateModified = System.DateTime.FromFileTime(ftdm)

    '            End If

    '            '                ListBox1.Items.Insert(i, filename.ToString() & " size:" & size & " date:" & DateModified.ToString())
    '            ListBox1.Items.Insert(i, System.Runtime.InteropServices.Marshal.PtrToStringUni(Everything_GetResultFileNameW(i)) & " Size:" & size & " Date Modified:" & DateModified.ToString())
    '        Next
    '    End If

    'End Sub

    Friend Function EverythingSearch(ByVal criteria As String) As String

        Dim ReturnValue As String = String.Empty

        Dim CurrentDirectory As String = Environment.CurrentDirectory

        Try

            Dim i As UInt32 = 0
            Dim numberOfResults As UInt32
            Dim filename As New System.Text.StringBuilder(260)
            Dim filenames As New List(Of String)

            If Environment.Is64BitProcess Then

                Everything_64Bit_SetSearchW(criteria)
                Everything_64Bit_SetRequestFlags(EVERYTHING_REQUEST_FILE_NAME Or EVERYTHING_REQUEST_PATH)
                Everything_64Bit_QueryW(1)
                numberOfResults = Everything_64Bit_GetNumResults()

                If numberOfResults > 0 Then

                    For x As UInt32 = 0 To numberOfResults - 1
                        Everything_64Bit_GetResultFullPathNameW(x, filename, filename.Capacity)
                        filenames.Add(filename.ToString)
                    Next

                End If

            Else

                If CurrentDirectory.ToLower.Contains("x86") Then

                    Everything_32Bit_x86_SetSearchW(criteria)
                    Everything_32Bit_x86_SetRequestFlags(EVERYTHING_REQUEST_FILE_NAME Or EVERYTHING_REQUEST_PATH)
                    Everything_32Bit_x86_QueryW(1)
                    numberOfResults = Everything_32Bit_x86_GetNumResults()

                    If numberOfResults > 0 Then

                        For x As UInt32 = 0 To numberOfResults - 1
                            Everything_32Bit_x86_GetResultFullPathNameW(x, filename, filename.Capacity)
                            filenames.Add(filename.ToString)
                        Next

                    End If

                Else

                    Everything_32Bit_SetSearchW(criteria)
                    Everything_32Bit_SetRequestFlags(EVERYTHING_REQUEST_FILE_NAME Or EVERYTHING_REQUEST_PATH)
                    Everything_32Bit_QueryW(1)
                    numberOfResults = Everything_32Bit_GetNumResults()

                    If numberOfResults > 0 Then

                        For x As UInt32 = 0 To numberOfResults - 1
                            Everything_32Bit_GetResultFullPathNameW(x, filename, filename.Capacity)
                            filenames.Add(filename.ToString)
                        Next

                    End If

                End If

            End If

            If numberOfResults > 0 Then

                filenames.Sort()
                ReturnValue = "" & filenames.First & ""

                If numberOfResults = 1 Then
                    Log("Everything matched on one file.")
                    Log("It was: " & ReturnValue)
                Else
                    Log("Everything matched on " & numberOfResults & " files.")
                    Log("The first match will be used, it was: " & ReturnValue)
                End If

            Else

                Log("Everything did not find a matching file.")

            End If

        Catch ex As Exception

            Log("Everything search failed." & vbCrLf & CurrentDirectory & vbCrLf & ex.Message.ToString)

        End Try

        Return ReturnValue

    End Function

End Module
