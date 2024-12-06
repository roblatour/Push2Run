Module modCommon
    Public Class myFormLibrary

        Public Shared WindowMain As Object

    End Class

    Friend gCapsLock As Boolean
    Friend gNumbLock As Boolean
    Friend gScrollLock As Boolean

    Friend gHandle As IntPtr
    Friend Sub GetLockKeyStates()

        myFormLibrary.WindowMain.Safely_GetKeylockStates()
        Application.DoEvents()

    End Sub

    Friend Sub MyDoEvents()

        Application.DoEvents()

    End Sub

    Friend gSendingKeyesIsRequired As Boolean

    Friend Function CaptureWindow(ByVal Handle As IntPtr) As System.Drawing.Image
        Return Nothing
        'dummy
    End Function

    Friend Function CaptureScreen() As System.Drawing.Image
        Return Nothing
        'dummy
    End Function

    Friend Function GetActiveWindowHandle() As Integer
        Return 0
    End Function


End Module
