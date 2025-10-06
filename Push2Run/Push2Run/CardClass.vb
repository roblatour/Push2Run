Imports System.Xml.Serialization ' Does XML serializing for a class.

' Set this 'Card' class as the root node of any XML file its serialized to.
Public Class CardClass

    Implements IDisposable

    Public Sub New()
    End Sub

    ' Set this 'DateTimeValue' field to be an attribute of the root node.
    ' Note date type must be "dateTime" in that odd mixture of upper and lower case to work
    <XmlAttributeAttribute(DataType:="dateTime")>
    Public DateTimeValue As System.DateTime = Now

    Public Description As String
    Public ListenFor As String
    Public Open As String
    Public StartDirectory As String
    Public Parameters As String
    Public StartWithAdminPrivileges As Boolean
    Public StartingWindowState As Integer
    Public KeysToSend As String

#Region "IDisposable Support"

    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' nothing to do really
            End If
        End If
        Me.disposedValue = True
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

#End Region

End Class

