Imports System.Runtime.CompilerServices
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
            If boolDoEscaping = True Then needle = Regex.Escape(needle)
            Return Regex.IsMatch(haystack, needle, RegexOptions.IgnoreCase)
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function doesPIDExist(PID As Integer) As Boolean
        Try
            Dim searcher As New Management.ManagementObjectSearcher("root\CIMV2", String.Format("Select * FROM Win32_Process WHERE ProcessId={0}", PID))

            If searcher.Get.Count = 0 Then
                searcher.Dispose()
                Return False
            Else
                searcher.Dispose()
                Return True
            End If
        Catch ex3 As Runtime.InteropServices.COMException
            Return False
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub killProcess(PID As Integer)
        Debug.Write(String.Format("Killing PID {0}...", PID))

        If doesPIDExist(PID) Then
            Process.GetProcessById(PID).Kill()
        End If

        If doesPIDExist(PID) Then
            killProcess(PID)
            'Else
            'debug.writeline(" Process Killed.")
        End If
    End Sub

    Public Sub searchForProcessAndKillIt(strFileName As String, boolFullFilePathPassed As Boolean)
        Dim fullFileName As String

        If boolFullFilePathPassed = True Then
            fullFileName = strFileName
        Else
            fullFileName = New IO.FileInfo(strFileName).FullName
        End If

        Dim wmiQuery As String = String.Format("Select ExecutablePath, ProcessId FROM Win32_Process WHERE ExecutablePath = '{0}'", fullFileName.addSlashes())
        Dim searcher As New Management.ManagementObjectSearcher("root\CIMV2", wmiQuery)

        Try
            For Each queryObj As Management.ManagementObject In searcher.Get()
                killProcess(Integer.Parse(queryObj("ProcessId").ToString))
            Next

            'debug.writeline("All processes killed... Update process can continue.")
        Catch ex3 As Runtime.InteropServices.COMException
        Catch err As Management.ManagementException
            ' Does nothing
        End Try
    End Sub

    Public Function areWeAnAdministrator() As Boolean
        Try
            Dim principal As WindowsPrincipal = New WindowsPrincipal(WindowsIdentity.GetCurrent())

            If principal.IsInRole(WindowsBuiltInRole.Administrator) = True Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function canIWriteToTheCurrentDirectory() As Boolean
        Return canIWriteThere(New IO.FileInfo(Application.ExecutablePath).DirectoryName)
    End Function

    Private Function canIWriteThere(folderPath As String) As Boolean
        ' We make sure we get valid folder path by taking off the leading slash.
        If folderPath.EndsWith("\") Then
            folderPath = folderPath.Substring(0, folderPath.Length - 1)
        End If

        If String.IsNullOrEmpty(folderPath) = True Or IO.Directory.Exists(folderPath) = False Then
            Return False
        End If

        If checkByFolderACLs(folderPath) = True Then
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
        Try
            Dim directoryACLs As DirectorySecurity = IO.Directory.GetAccessControl(folderPath)
            Dim directoryUsers As String = WindowsIdentity.GetCurrent.User.Value
            Dim directoryAccessRights As FileSystemAccessRule
            Dim fileSystemRights As FileSystemRights

            For Each rule As AuthorizationRule In directoryACLs.GetAccessRules(True, True, GetType(SecurityIdentifier))
                If rule.IdentityReference.Value = directoryUsers Then
                    directoryAccessRights = DirectCast(rule, FileSystemAccessRule)

                    If directoryAccessRights.AccessControlType = Security.AccessControl.AccessControlType.Allow Then
                        fileSystemRights = directoryAccessRights.FileSystemRights

                        If fileSystemRights = (FileSystemRights.Read Or FileSystemRights.Modify Or FileSystemRights.Write Or FileSystemRights.FullControl) Then
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