Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Security.AccessControl
Imports System.Security.Principal
Imports System.Text.RegularExpressions

Module Module1
    ' PHP like addSlashes and stripSlashes. Call using String.addSlashes() and String.stripSlashes().
    <Extension()>
    Public Function addSlashes(unsafeString As String) As String
        Return Regex.Replace(unsafeString, "([\000\010\011\012\015\032\042\047\134\140])", "\$1")
    End Function

    <Extension()>
    Public Function caseInsensitiveContains(haystack As String, needle As String, Optional boolDoEscaping As Boolean = False) As Boolean
        Try
            If boolDoEscaping Then needle = Regex.Escape(needle)
            Return Regex.IsMatch(haystack, needle, RegexOptions.IgnoreCase)
        Catch ex As Exception
            Return False
        End Try
    End Function

    <Extension()>
    Public Function stringCompare(str1 As String, str2 As String, Optional boolCaseInsensitive As Boolean = True)
        If boolCaseInsensitive Then
            Return str1.Trim.Equals(str2.Trim, StringComparison.OrdinalIgnoreCase)
        Else
            Return str1.Trim.Equals(str2.Trim, StringComparison.Ordinal)
        End If
    End Function

    ''' <summary>Checks to see if a Process ID or PID exists on the system.</summary>
    ''' <param name="PID">The PID of the process you are checking the existance of.</param>
    ''' <param name="processObject">If the PID does exist, the function writes back to this argument in a ByRef way a Process Object that can be interacted with outside of this function.</param>
    ''' <returns>Return a Boolean value. If the PID exists, it return a True value. If the PID doesn't exist, it returns a False value.</returns>
    Private Function doesPIDExist(ByVal PID As Integer, ByRef processObject As Process) As Boolean
        Try
            processObject = Process.GetProcessById(PID)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub killProcess(PID As Integer)
        Dim processObject As Process = Nothing
        If doesPIDExist(PID, processObject) Then
            Try
                processObject.Kill() ' Yes, it does so let's kill it.
            Catch ex As Exception
                ' Wow, it seems that even with double-checking if a process exists by it's PID number things can still go wrong.
                ' So this Try-Catch block is here to trap any possible errors when trying to kill a process by it's PID number.
            End Try
        End If
    End Sub

    Private Function getProcessExecutablePath(processID As Integer) As String
        Dim memoryBuffer = New Text.StringBuilder(1024)
        Dim processHandle As IntPtr = NativeMethods.OpenProcess(NativeMethods.ProcessAccessFlags.PROCESS_QUERY_LIMITED_INFORMATION, False, processID)

        If processHandle <> IntPtr.Zero Then
            Try
                Dim memoryBufferSize As Integer = memoryBuffer.Capacity

                If NativeMethods.QueryFullProcessImageName(processHandle, 0, memoryBuffer, memoryBufferSize) Then
                    Return memoryBuffer.ToString()
                End If
            Finally
                NativeMethods.CloseHandle(processHandle)
            End Try
        End If

        Return Nothing
    End Function

    Private Sub searchForProcessAndKillIt(strFileName As String, boolFullFilePathPassed As Boolean)
        Dim processExecutablePath As String
        Dim processExecutablePathFileInfo As IO.FileInfo

        For Each process As Process In Process.GetProcesses()
            processExecutablePath = getProcessExecutablePath(process.Id)

            If processExecutablePath IsNot Nothing Then
                processExecutablePathFileInfo = New IO.FileInfo(processExecutablePath)

                If boolFullFilePathPassed Then
                    If stringCompare(strFileName, processExecutablePathFileInfo.FullName) Then killProcess(process.Id)
                ElseIf Not boolFullFilePathPassed Then
                    If stringCompare(strFileName, processExecutablePathFileInfo.Name) Then killProcess(process.Id)
                End If

                processExecutablePathFileInfo = Nothing
            End If
        Next
    End Sub

    Public Function areWeAnAdministrator() As Boolean
        Try
            Dim principal As WindowsPrincipal = New WindowsPrincipal(WindowsIdentity.GetCurrent())
            Return principal.IsInRole(WindowsBuiltInRole.Administrator)
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function canIWriteToTheCurrentDirectory() As Boolean
        Return canIWriteThere(New IO.FileInfo(Application.ExecutablePath).DirectoryName)
    End Function

    Private Function canIWriteThere(folderPath As String) As Boolean
        ' We make sure we get valid folder path by taking off the leading slash.
        If folderPath.EndsWith("\") Then folderPath = folderPath.Substring(0, folderPath.Length - 1)
        If String.IsNullOrEmpty(folderPath) Or Not IO.Directory.Exists(folderPath) Then Return False

        If checkByFolderACLs(folderPath) Then
            Try
                IO.File.Create(IO.Path.Combine(folderPath, "test.txt"), 1, IO.FileOptions.DeleteOnClose).Close()
                If IO.File.Exists(IO.Path.Combine(folderPath, "test.txt")) Then IO.File.Delete(IO.Path.Combine(folderPath, "test.txt"))
                Return True
            Catch ex As Exception
                Return False
            End Try
        Else
            Return False
        End If
    End Function

    Private Function checkByFolderACLs(folderPath As String) As Boolean
        If WindowsIdentity.GetCurrent().IsSystem Then Return True

        Try
            Dim dsDirectoryACLs As DirectorySecurity = IO.Directory.GetAccessControl(folderPath)
            Dim strCurrentUserSDDL As String = WindowsIdentity.GetCurrent.User.Value
            Dim ircCurrentUserGroups As IdentityReferenceCollection = WindowsIdentity.GetCurrent.Groups

            Dim arcAuthorizationRules As AuthorizationRuleCollection = dsDirectoryACLs.GetAccessRules(True, True, GetType(SecurityIdentifier))
            Dim fsarDirectoryAccessRights As FileSystemAccessRule

            For Each arAccessRule As AuthorizationRule In arcAuthorizationRules
                If arAccessRule.IdentityReference.Value.Equals(strCurrentUserSDDL, StringComparison.OrdinalIgnoreCase) Or ircCurrentUserGroups.Contains(arAccessRule.IdentityReference) Then
                    fsarDirectoryAccessRights = DirectCast(arAccessRule, FileSystemAccessRule)

                    If fsarDirectoryAccessRights.AccessControlType = AccessControlType.Allow Then
                        If fsarDirectoryAccessRights.FileSystemRights = FileSystemRights.Modify Or fsarDirectoryAccessRights.FileSystemRights = FileSystemRights.WriteData Or fsarDirectoryAccessRights.FileSystemRights = FileSystemRights.FullControl Then
                            Return True
                        End If
                    End If
                End If
            Next

            Return False
        Catch ex As Exception
            Return False
        End Try
    End Function
End Module

Friend NotInheritable Class NativeMethods
    Private Sub New()
    End Sub

    <Flags>
    Public Enum ProcessAccessFlags As UInteger
        PROCESS_QUERY_LIMITED_INFORMATION = &H1000
        All = &H1F0FFF
        Terminate = &H1
        CreateThread = &H2
        VirtualMemoryOperation = &H8
        VirtualMemoryRead = &H10
        VirtualMemoryWrite = &H20
        DuplicateHandle = &H40
        CreateProcess = &H80
        SetQuota = &H100
        SetInformation = &H200
        QueryInformation = &H400
        QueryLimitedInformation = &H1000
        Synchronize = &H100000
    End Enum

    <DllImport("kernel32.dll", CharSet:=CharSet.Unicode)>
    Friend Shared Function QueryFullProcessImageName(hprocess As IntPtr, dwFlags As Integer, lpExeName As Text.StringBuilder, ByRef size As Integer) As Boolean
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Unicode)>
    Friend Shared Function OpenProcess(dwDesiredAccess As ProcessAccessFlags, bInheritHandle As Boolean, dwProcessId As Integer) As IntPtr
    End Function

    <DllImport("kernel32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Friend Shared Function CloseHandle(hHandle As IntPtr) As Boolean
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Unicode)>
    Friend Shared Function MoveFileEx(ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal dwFlags As Int32) As Boolean
    End Function
End Class