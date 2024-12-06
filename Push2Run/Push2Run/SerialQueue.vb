'ref https://github.com/Gentlee/SerialQueue

Namespace Threading
    Public Class SerialQueue

        Private ReadOnly _locker As Object = New Object()
        Private ReadOnly _lastTask As WeakReference(Of Task) = New WeakReference(Of Task)(Nothing)

        Public Function Enqueue(ByVal action As Action) As Task
            Return Enqueue(Function()
                               action()
                               Return True
                           End Function)
        End Function

        Public Function Enqueue(Of T)(ByVal [function] As Func(Of T)) As Task(Of T)
            SyncLock _locker
                Dim lastTask As Task
                Dim resultTask As Task(Of T)

                If _lastTask.TryGetTarget(lastTask) Then
                    resultTask = lastTask.ContinueWith(Function(__) [function](), TaskContinuationOptions.ExecuteSynchronously)
                Else
                    resultTask = Task.Run([function])
                End If

                _lastTask.SetTarget(resultTask)
                Return resultTask
            End SyncLock
        End Function

        Public Function Enqueue(ByVal asyncAction As Func(Of Task)) As Task
            SyncLock _locker
                Dim lastTask As Task
                Dim resultTask As Task

                If _lastTask.TryGetTarget(lastTask) Then
                    resultTask = lastTask.ContinueWith(Function(__) asyncAction(), TaskContinuationOptions.ExecuteSynchronously).Unwrap()
                Else
                    resultTask = Task.Run(asyncAction)
                End If

                _lastTask.SetTarget(resultTask)
                Return resultTask
            End SyncLock
        End Function

        Public Function Enqueue(Of T)(ByVal asyncFunction As Func(Of Task(Of T))) As Task(Of T)
            SyncLock _locker
                Dim lastTask As Task
                Dim resultTask As Task(Of T)

                If _lastTask.TryGetTarget(lastTask) Then
                    resultTask = lastTask.ContinueWith(Function(__) asyncFunction(), TaskContinuationOptions.ExecuteSynchronously).Unwrap()
                Else
                    resultTask = Task.Run(asyncFunction)
                End If

                _lastTask.SetTarget(resultTask)
                Return resultTask
            End SyncLock
        End Function
    End Class
End Namespace

