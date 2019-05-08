Imports Microsoft.Win32
Imports System.Security.Principal
Imports System.Text.RegularExpressions
Imports System.Runtime.InteropServices
Imports System.Security.AccessControl
Imports System.Runtime.CompilerServices

Namespace programConstants
    Module programConstants
        Public Const errorRetrievingRemoteINIFileVersion As String = "Error Retrieving Remote INI File Version"
        Public Const customEntriesFile As String = "YAWA2 Updater Custom Entries.txt"
        Public Const configINIFile As String = "YAWA2 Updater Config.ini"

        Public Const WinApp2INIFileURL As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Winapp2.ini"

        Public Const configINISettingSection As String = "Configuration"
        Public Const configINICustomEntriesKey As String = "customEntries"
        Public Const configINIMobileModeKey As String = "MobileMode"
        Public Const configINITrimKey As String = "trim"
        Public Const configINInotifyAfterUpdateAtLogonKey As String = "notifyAfterUpdateAtLogon"
    End Module
End Namespace

Module globalVariables
    Public boolUseSSL As Boolean = True
End Module

Namespace programVariables
    Module variables
        Public boolMobileMode, boolTrim, boolNotifyAfterUpdateAtLogon As Boolean
    End Module
End Namespace

Namespace programFunctions
    Module functions
        Public ReadOnly osVersionString As String = Environment.OSVersion.Version.Major & "." & Environment.OSVersion.Version.Minor

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

            NativeMethods.CloseHandle(processHandle)
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

        Public Function areWeAnAdministrator() As Boolean
            Try
                Dim principal As New WindowsPrincipal(WindowsIdentity.GetCurrent())
                Return principal.IsInRole(WindowsBuiltInRole.Administrator)
            Catch ex As Exception
                Return False
            End Try
        End Function

        Public Function isBase64(base64String As String) As Boolean
            If base64String.Replace(" ", "").Length Mod 4 <> 0 Then Return False
            Try
                Convert.FromBase64String(base64String)
                Return True
            Catch exception As FormatException ' Handle the exception
                Return False
            End Try
        End Function

        Public Function convertToBase64(input As String) As String
            Return Convert.ToBase64String(Text.Encoding.UTF8.GetBytes(input))
        End Function

        Public Function convertFromBase64(input As String) As String
            Return Text.Encoding.UTF8.GetString(Convert.FromBase64String(input))
        End Function

        Public Sub removeSettingFromINIFile(settingToRemove As String)
            ' Check if the INI file exists.
            If IO.File.Exists(programConstants.configINIFile) Then
                ' Yes, it does exist; now let's load the file into memory.

                Dim iniFile As New IniFile() ' First we create an IniFile class object.
                iniFile.loadINIFileFromFile(programConstants.configINIFile) ' Now we load the data into memory.

                Dim temp As String = iniFile.GetKeyValue(programConstants.configINISettingSection, settingToRemove) ' Load the data from the INI object into local program memory.
                If String.IsNullOrWhiteSpace(temp) Then
                    iniFile.RemoveKey(programConstants.configINISettingSection, settingToRemove) ' Remove the setting from the INI file in memory.
                End If

                iniFile.Save(programConstants.configINIFile) ' Save the data to disk.
                iniFile = Nothing ' And destroy the INIFile object.
            End If
        End Sub

        Public Sub saveSettingToINIFile(setting As String, value As Boolean)
            saveSettingToINIFile(setting, If(value, "True", "False"))
        End Sub

        Public Sub saveSettingToINIFile(setting As String, value As String)
            ' Check if the INI file exists.
            If IO.File.Exists(programConstants.configINIFile) Then
                ' Yes, it does exist; now let's load the file into memory.

                Dim iniFile As New IniFile() ' First we create an IniFile class object.

                ' Now, we have to load the existing data into memory or we're going to end up overwriting the existing INI file with just
                ' the setting we're setting now. Essentially you'd end up with an INI file with only one entry in it and that's bad.
                iniFile.loadINIFileFromFile(programConstants.configINIFile)

                iniFile.SetKeyValue(programConstants.configINISettingSection, setting, value) ' We now set the value.
                iniFile.Save(programConstants.configINIFile) ' Save what's in memory to disk.
                iniFile = Nothing ' And destroy the INIFile object.
            Else
                ' No, it doesn't exist so we just create a new INI object with no prior data to load.

                Dim iniFile As New IniFile() ' First we create an IniFile class object.
                iniFile.SetKeyValue(programConstants.configINISettingSection, setting, value) ' We now set the value.
                iniFile.Save(programConstants.configINIFile) ' Save what's in memory to disk.
                iniFile = Nothing ' And destroy the INIFile object.
            End If
        End Sub

        Public Function loadSettingFromINIFile(ByVal settingKey As String, ByRef settingKeyValue As String) As Boolean
            Try
                Dim iniFile As New IniFile() ' First we create an IniFile class object.

                ' Check if the INI file exists.
                If IO.File.Exists(programConstants.configINIFile) Then
                    iniFile.loadINIFileFromFile(programConstants.configINIFile) ' Yes, it does exist; now let's load the file into memory.

                    ' Load the data from the INI object into local program memory.
                    Dim temp As String = iniFile.GetKeyValue(programConstants.configINISettingSection, settingKey)
                    iniFile = Nothing ' And destroy the INIFile object.

                    If String.IsNullOrWhiteSpace(temp) Then
                        Return False ' OK, the setting in the INI file doesn't exist so we return with a False Boolean value.
                    Else
                        settingKeyValue = temp ' Put the value into the ByRef settingKeyValue variable.
                        Return True ' And return with a True Boolean value.
                    End If
                Else
                    ' No, it doesn't exist so we simply return with a False Boolean value.
                    Return False
                End If
            Catch ex As Exception
                Debug.WriteLine(ex.Message & " -- " & ex.StackTrace.Trim)
                Return False
            End Try
        End Function

        Public Function getRemoteINIFileVersion() As String
            Try
                Dim httpHelper As httpHelper = internetFunctions.createNewHTTPHelperObject()
                Dim strINIFileData As String = Nothing
                Return If(httpHelper.getWebData(programConstants.WinApp2INIFileURL, strINIFileData, False), getINIVersionFromString(strINIFileData), programConstants.errorRetrievingRemoteINIFileVersion)
            Catch ex As Exception
                Return programConstants.errorRetrievingRemoteINIFileVersion
            End Try
        End Function

        Private Function osVersionCheck(tempString As String) As Boolean
            If Environment.OSVersion.Version.Major = 10 Then
                If tempString.Contains("6.2") Or tempString.Contains("10.0") Then Return True
            Else
                If tempString.Contains(osVersionString) Then Return True
            End If
            Return False
        End Function

        Public Sub trimINIFile(strLocationOfCCleaner As String, remoteINIFileVersion As String, boolSilentMode As Boolean)
            Dim tempString As String
            Dim oldINIFileContents As String
            Dim sectionsToRemove As New Specialized.StringCollection

            Using streamReader As New IO.StreamReader(IO.Path.Combine(strLocationOfCCleaner, "winapp2.ini"))
                oldINIFileContents = streamReader.ReadToEnd.Replace(vbLf, vbCrLf).Replace(vbCr, Nothing).Replace(vbLf, vbCrLf)
            End Using

            Dim matchData As Match = Regex.Match(oldINIFileContents, "(; # of entries: ([0-9,]*).*; Chrome/Chromium based browsers)", RegexOptions.Singleline Or RegexOptions.IgnoreCase)
            Dim iniFileNotes As String = matchData.Groups(1).Value()
            Dim entriesString As String = matchData.Groups(2).Value()
            matchData = Nothing

            Dim iniFile As New IniFile()
            iniFile.loadINIFileFromFile(IO.Path.Combine(strLocationOfCCleaner, "winapp2.ini"))

            For Each iniFileSection As IniFile.IniSection In iniFile.Sections
                tempString = iniFile.GetKeyValue(iniFileSection.Name, "DetectOS")
                If Not String.IsNullOrWhiteSpace(tempString) Then
                    If Not osVersionCheck(tempString) Then
                        If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                            sectionsToRemove.Add(iniFileSection.Name)
                            Continue For
                        End If
                    End If
                End If
                tempString = Nothing

                tempString = iniFile.GetKeyValue(iniFileSection.Name, "Detect")
                If Not String.IsNullOrWhiteSpace(tempString) Then
                    If processRegistryKey(tempString, sectionsToRemove, iniFileSection) Then
                        Continue For
                    End If
                End If
                tempString = Nothing

                tempString = iniFile.GetKeyValue(iniFileSection.Name, "DetectFile")
                If Not String.IsNullOrWhiteSpace(tempString) Then
                    If processFilePath(tempString, sectionsToRemove, iniFileSection) Then
                        Continue For
                    End If
                End If
                tempString = Nothing

                tempString = iniFile.GetKeyValue(iniFileSection.Name, "Detect1")
                If Not String.IsNullOrWhiteSpace(tempString) Then
                    If processRegistryKey(tempString, sectionsToRemove, iniFileSection) Then
                        Continue For
                    End If
                End If
                tempString = Nothing

                tempString = iniFile.GetKeyValue(iniFileSection.Name, "DetectFile1")
                If Not String.IsNullOrWhiteSpace(tempString) Then
                    If processFilePath(tempString, sectionsToRemove, iniFileSection) Then
                        Continue For
                    End If
                End If
                tempString = Nothing

                tempString = iniFile.GetKeyValue(iniFileSection.Name, "FileKey1")
                If Not String.IsNullOrWhiteSpace(tempString) Then
                    If processFilePath(tempString, sectionsToRemove, iniFileSection) Then
                        Continue For
                    End If
                End If
                tempString = Nothing

                tempString = iniFile.GetKeyValue(iniFileSection.Name, "RegKey1")
                If Not String.IsNullOrWhiteSpace(tempString) Then
                    If processRegistryKey(tempString, sectionsToRemove, iniFileSection) Then
                        Continue For
                    End If
                End If
                tempString = Nothing
            Next

            For Each sectionToRemove As String In sectionsToRemove
                iniFile.RemoveSection(sectionToRemove)
            Next

            Dim rawINIFileContents As String = iniFile.getRawINIText

            iniFileNotes = iniFileNotes.Replace(entriesString, iniFile.Sections.Count.ToString("N0"))
            iniFile = Nothing

            Dim newINIFileContents As String = "; Version: " & remoteINIFileVersion & vbCrLf
            newINIFileContents &= "; Last Updated On: " & Now.Date.ToLongDateString & vbCrLf
            newINIFileContents &= iniFileNotes & vbCrLf & vbCrLf
            newINIFileContents &= rawINIFileContents

            If IO.File.Exists(IO.Path.Combine(strLocationOfCCleaner, "winapp2.ini")) Then
                IO.File.Delete(IO.Path.Combine(strLocationOfCCleaner, "winapp2.ini"))
            End If

            Using streamWriter As New IO.StreamWriter(IO.Path.Combine(strLocationOfCCleaner, "winapp2.ini"))
                streamWriter.Write(newINIFileContents)
            End Using

            ' These are potentially large variables, we need to trash them to free up memory.
            oldINIFileContents = Nothing
            newINIFileContents = Nothing
            rawINIFileContents = Nothing

            If Not boolSilentMode Then
                If programVariables.boolMobileMode Then : MsgBox("INI File Trim Complete.  A total of " & sectionsToRemove.Count.ToString("N0", Globalization.CultureInfo.CreateSpecificCulture("en-US")) & " sections were removed.", MsgBoxStyle.Information, "WinApp.ini Updater")
                Else
                    Dim msgBoxResult As MsgBoxResult = MsgBox("INI File Trim Complete.  A total of " & sectionsToRemove.Count.ToString("N0", Globalization.CultureInfo.CreateSpecificCulture("en-US")) & " sections were removed." & vbCrLf & vbCrLf & "Do you want to run CCleaner now?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "WinApp.ini Updater")

                    If msgBoxResult = Microsoft.VisualBasic.MsgBoxResult.Yes Then
                        runCCleaner(strLocationOfCCleaner)
                    End If
                End If
            End If
        End Sub

        Public Sub runCCleaner(strLocationOfCCleaner As String)
            Process.Start(IO.Path.Combine(strLocationOfCCleaner, If(Environment.Is64BitOperatingSystem, "CCleaner64.exe", "CCleaner.exe")))
        End Sub

        Private Function translateVarsInPath(input As String) As String
            input = input.Replace("%AppData%", Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData))
            input = input.Replace("%LocalAppData%", Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData))
            input = input.Replace("%Documents%", Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments))
            input = input.Replace("%CommonAppData%", Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData))
            input = input.Replace("%SystemDrive%", Environment.GetFolderPath(Environment.SpecialFolder.Windows).Replace("\Windows", ""))
            input = input.Replace("%WinDir%", Environment.GetFolderPath(Environment.SpecialFolder.Windows))
            Return input
        End Function

        Public Function getINIVersionFromString(input As String) As String
            ' Special Regular Expression to extract the version of the INI file from the INI file's raw text.
            Return Regex.Match(input, "; Version: ([0-9.A-Za-z]+)").Groups(1).Value
        End Function

        Public Function processFilePath(ByVal tempString As String, ByRef sectionsToRemove As Specialized.StringCollection, ByRef iniFileSection As IniFile.IniSection) As Boolean
            Dim directory As String = tempString.Split("|")(0)
            directory = directory.Replace("*", "")

            If directory.Contains("%ProgramFiles%") Then
                If Environment.Is64BitOperatingSystem Then
                    If Not IO.Directory.Exists(directory.Replace("%ProgramFiles%", Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles))) And Not IO.Directory.Exists(directory.Replace("%ProgramFiles%", Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86))) Then
                        If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                            sectionsToRemove.Add(iniFileSection.Name)
                            Return True
                        Else : Return True
                        End If
                    Else : Return True
                    End If
                Else
                    If Not IO.Directory.Exists(directory.Replace("%ProgramFiles%", Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles))) Then
                        If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                            sectionsToRemove.Add(iniFileSection.Name)
                            Return True
                        Else : Return True
                        End If
                    Else : Return True
                    End If
                End If
            ElseIf directory.Contains("%CommonProgramFiles%") Then
                If Environment.Is64BitOperatingSystem Then
                    If Not IO.Directory.Exists(directory.Replace("%CommonProgramFiles%", Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles))) And Not IO.Directory.Exists(directory.Replace("%CommonProgramFiles%", Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFilesX86))) Then
                        If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                            sectionsToRemove.Add(iniFileSection.Name)
                            Return True
                        Else : Return True
                        End If
                    Else : Return True
                    End If
                Else
                    If Not IO.Directory.Exists(directory.Replace("%CommonProgramFiles%", Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles))) Then
                        If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                            sectionsToRemove.Add(iniFileSection.Name)
                            Return True
                        Else : Return True
                        End If
                    Else : Return True
                    End If
                End If
            Else
                directory = translateVarsInPath(directory)
                directory = directory.Replace("*", "")

                If Not IO.Directory.Exists(directory) Then
                    If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                        sectionsToRemove.Add(iniFileSection.Name)
                        Return True
                    Else : Return True
                    End If
                Else : Return True
                End If
            End If

            Return False
        End Function

        Public Function processRegistryKey(ByVal tempString As String, ByRef sectionsToRemove As Specialized.StringCollection, ByRef iniFileSection As IniFile.IniSection) As Boolean
            Try
                If tempString.Contains(".NETFramework") Then Return True

                If tempString.StartsWith("HKCU") Then
                    tempString = tempString.Replace("HKCU\", "")

                    If Environment.Is64BitOperatingSystem Then
                        If Boolean.Parse(RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry32).OpenSubKey(tempString) Is Nothing) And Boolean.Parse(RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64).OpenSubKey(tempString) Is Nothing) Then
                            If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                                sectionsToRemove.Add(iniFileSection.Name)
                                Return True
                            Else : Return True
                            End If
                        Else : Return True
                        End If
                    Else
                        If Boolean.Parse(Registry.CurrentUser.OpenSubKey(tempString) Is Nothing) Then
                            If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                                sectionsToRemove.Add(iniFileSection.Name)
                                Return True
                            Else : Return True
                            End If
                        Else : Return True
                        End If
                    End If
                ElseIf tempString.StartsWith("HKLM") Then
                    tempString = tempString.Replace("HKLM\", "")

                    If Environment.Is64BitOperatingSystem Then
                        If Boolean.Parse(RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32).OpenSubKey(tempString) Is Nothing) And Boolean.Parse(RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(tempString) Is Nothing) Then
                            If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                                sectionsToRemove.Add(iniFileSection.Name)
                                Return True
                            Else : Return True
                            End If
                        Else : Return True
                        End If
                    Else
                        If Boolean.Parse(Registry.LocalMachine.OpenSubKey(tempString) Is Nothing) Then
                            If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                                sectionsToRemove.Add(iniFileSection.Name)
                                Return True
                            Else : Return True
                            End If
                        Else : Return True
                        End If
                    End If
                ElseIf tempString.StartsWith("HKCR") Then
                    tempString = tempString.Replace("HKCR\", "")

                    If Environment.Is64BitOperatingSystem Then
                        If Boolean.Parse(RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry32).OpenSubKey(tempString) Is Nothing) And Boolean.Parse(RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry64).OpenSubKey(tempString) Is Nothing) Then
                            If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                                sectionsToRemove.Add(iniFileSection.Name)
                                Return True
                            Else : Return True
                            End If
                        Else : Return True
                        End If
                    Else
                        If Boolean.Parse(Registry.ClassesRoot.OpenSubKey(tempString) Is Nothing) Then
                            If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                                sectionsToRemove.Add(iniFileSection.Name)
                                Return True
                            Else : Return True
                            End If
                        Else : Return True
                        End If
                    End If
                ElseIf tempString.StartsWith("HKU") Then
                    tempString = tempString.Replace("HKU\", "")

                    If Environment.Is64BitOperatingSystem Then
                        If Boolean.Parse(RegistryKey.OpenBaseKey(RegistryHive.Users, RegistryView.Registry32).OpenSubKey(tempString) Is Nothing) And Boolean.Parse(RegistryKey.OpenBaseKey(RegistryHive.Users, RegistryView.Registry64).OpenSubKey(tempString) Is Nothing) Then
                            If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                                sectionsToRemove.Add(iniFileSection.Name)
                                Return True
                            Else : Return True
                            End If
                        Else : Return True
                        End If
                    Else
                        If Boolean.Parse(Registry.Users.OpenSubKey(tempString) Is Nothing) Then
                            If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                                sectionsToRemove.Add(iniFileSection.Name)
                                Return True
                            Else : Return True
                            End If
                        Else : Return True
                        End If
                    End If
                End If

                Return False
            Catch ex As Security.SecurityException
                Return False
            End Try
        End Function
    End Module
End Namespace

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