Module Search_for_Process_and_Kill_it
    ''' <summary>Checks to see if a Process ID or PID exists on the system.</summary>
    ''' <param name="PID">The PID of the process you are checking the existance of.</param>
    ''' <param name="processObject">If the PID does exist, the function writes back to this argument in a ByRef way a Process Object that can be interacted with outside of this function.</param>
    ''' <returns>Return a Boolean value. If the PID exists, it return a True value. If the PID doesn't exist, it returns a False value.</returns>
    Private Function DoesProcessIDExist(PID As Integer, ByRef processObject As Process) As Boolean
        Try
            processObject = Process.GetProcessById(PID)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub KillProcess(processID As Integer)
        KillProcessSubRoutine(processID)
        Threading.Thread.Sleep(250) ' We're going to sleep to give the system some time to kill the process.
        KillProcessSubRoutine(processID)
        Threading.Thread.Sleep(250) ' We're going to sleep (again) to give the system some time to kill the process.
    End Sub

    Private Sub KillProcessSubRoutine(processID As Integer)
        Dim processObject As Process = Nothing
        If DoesProcessIDExist(processID, processObject) Then
            Try
                processObject.Kill()
            Catch ex As Exception
                ' Wow, it seems that even with double-checking if a process exists by it's PID number things can still go wrong.
                ' So this Try-Catch block is here to trap any possible errors when trying to kill a process by it's PID number.
            End Try
        End If
    End Sub

    Public Sub SearchForProcessAndKillIt(strFileName As String, boolFullFilePathPassed As Boolean)
        Dim processExecutablePath As String
        Dim processExecutablePathFileInfo As IO.FileInfo

        For Each process As Process In Process.GetProcesses()
            processExecutablePath = GetProcessExecutablePath(process.Id)

            If Not String.IsNullOrWhiteSpace(processExecutablePath) Then
                Try
                    processExecutablePathFileInfo = New IO.FileInfo(processExecutablePath)
                    processExecutablePath = If(boolFullFilePathPassed, processExecutablePathFileInfo.FullName, processExecutablePathFileInfo.Name)
                    If strFileName.Equals(processExecutablePath, StringComparison.OrdinalIgnoreCase) Then KillProcess(process.Id)
                Catch ex As ArgumentException
                End Try
            End If
        Next
    End Sub

    Private Function GetProcessExecutablePath(processID As Integer) As String
        Try
            Dim memoryBuffer As New Text.StringBuilder(1024)
            Dim processHandle As IntPtr = NativeMethod.NativeMethods.OpenProcess(NativeMethod.ProcessAccessFlags.PROCESS_QUERY_LIMITED_INFORMATION, False, processID)

            If Not processHandle.Equals(IntPtr.Zero) Then
                Try
                    Dim memoryBufferSize As Integer = memoryBuffer.Capacity
                    If NativeMethod.NativeMethods.QueryFullProcessImageName(processHandle, 0, memoryBuffer, memoryBufferSize) Then Return memoryBuffer.ToString()
                Finally
                    NativeMethod.NativeMethods.CloseHandle(processHandle)
                End Try
            End If

            NativeMethod.NativeMethods.CloseHandle(processHandle)
            Return Nothing
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
End Module