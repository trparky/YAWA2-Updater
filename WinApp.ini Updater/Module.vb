Imports Microsoft.Win32
Imports System.Security.Principal

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
    Public boolWinXP As Boolean = False
    Public boolUseSSL As Boolean = True
End Module

Namespace programVariables
    Module variables
        Public boolMobileMode, boolTrim, boolNotifyAfterUpdateAtLogon As Boolean
    End Module
End Namespace

Namespace programFunctions
    Module functions
        Public Function areWeAnAdministrator() As Boolean
            Try
                Dim principal As New WindowsPrincipal(WindowsIdentity.GetCurrent())

                If principal.IsInRole(WindowsBuiltInRole.Administrator) = True Then
                    Return (True)
                Else
                    Return False
                End If
            Catch ex As Exception
                Return False
            End Try
        End Function

        Public Function isBase64(base64String As String) As Boolean
            If base64String.Replace(" ", "").Length Mod 4 <> 0 Then
                Return False
            End If

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
                If temp Is Nothing Or temp = Nothing Then
                    iniFile.RemoveKey(programConstants.configINISettingSection, settingToRemove) ' Remove the setting from the INI file in memory.
                End If

                iniFile.Save(programConstants.configINIFile) ' Save the data to disk.
                iniFile = Nothing ' And destroy the INIFile object.
            End If
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

                    If temp Is Nothing Or temp = Nothing Then
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

                If httpHelper.getWebData(programConstants.WinApp2INIFileURL, strINIFileData, False) Then
                    Return getINIVersionFromString(strINIFileData)
                Else
                    Return programConstants.errorRetrievingRemoteINIFileVersion
                End If
            Catch ex As Exception
                Return programConstants.errorRetrievingRemoteINIFileVersion
            End Try
        End Function

        Public Sub trimINIFile(strLocationOfCCleaner As String, remoteINIFileVersion As String, boolSilentMode As Boolean)
            Dim tempString As String, streamReader As IO.StreamReader
            Dim oldINIFileContents As String
            Dim sectionsToRemove As New Specialized.StringCollection

            streamReader = New IO.StreamReader(IO.Path.Combine(strLocationOfCCleaner, "winapp2.ini"))
            oldINIFileContents = streamReader.ReadToEnd
            streamReader.Close()
            streamReader.Dispose()

            Dim matchData As Text.RegularExpressions.Match = Text.RegularExpressions.Regex.Match(oldINIFileContents, "(; Application Cleaning file" & vbCrLf & ";" & vbCrLf & "; Notes" & vbCrLf & "; # of entries: ([0-9,]*)" & vbCrLf & ".*" & vbCrLf & "; Please do not host this file anywhere without permission\. This is to facilitate proper distribution of the latest version\. Thanks\.)", System.Text.RegularExpressions.RegexOptions.Singleline Or Text.RegularExpressions.RegexOptions.IgnoreCase)
            Dim iniFileNotes As String = matchData.Groups(1).Value()
            Dim entriesString As String = matchData.Groups(2).Value()
            matchData = Nothing

            Dim iniFile As New IniFile()
            iniFile.loadINIFileFromFile(IO.Path.Combine(strLocationOfCCleaner, "winapp2.ini"))

            For Each iniFileSection As IniFile.IniSection In iniFile.Sections
                'Trace.WriteLine(String.Format("Section: [{0}]", iniFileSection.Name))

                tempString = iniFile.GetKeyValue(iniFileSection.Name, "DetectOS")
                If tempString <> "" Then
                    If Not tempString.Contains(Environment.OSVersion.Version.Major & "." & Environment.OSVersion.Version.Minor) Then
                        If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                            sectionsToRemove.Add(iniFileSection.Name)
                            Continue For
                        End If
                    End If
                End If
                tempString = ""

                tempString = iniFile.GetKeyValue(iniFileSection.Name, "Detect")
                If tempString <> "" Then
                    If processRegistryKey(tempString, sectionsToRemove, iniFileSection) Then
                        Continue For
                    End If
                End If
                tempString = ""

                tempString = iniFile.GetKeyValue(iniFileSection.Name, "DetectFile")
                If tempString <> "" Then
                    If processFilePath(tempString, sectionsToRemove, iniFileSection) Then
                        Continue For
                    End If
                End If
                tempString = ""

                tempString = iniFile.GetKeyValue(iniFileSection.Name, "Detect1")
                If tempString <> "" Then
                    If processRegistryKey(tempString, sectionsToRemove, iniFileSection) Then
                        Continue For
                    End If
                End If
                tempString = ""

                tempString = iniFile.GetKeyValue(iniFileSection.Name, "DetectFile1")
                If tempString <> "" Then
                    If processFilePath(tempString, sectionsToRemove, iniFileSection) Then
                        Continue For
                    End If
                End If
                tempString = ""

                tempString = iniFile.GetKeyValue(iniFileSection.Name, "FileKey1")
                If tempString <> "" Then
                    If processFilePath(tempString, sectionsToRemove, iniFileSection) Then
                        Continue For
                    End If
                End If
                tempString = ""

                tempString = iniFile.GetKeyValue(iniFileSection.Name, "RegKey1")
                If tempString <> "" Then
                    If processRegistryKey(tempString, sectionsToRemove, iniFileSection) Then
                        Continue For
                    End If
                End If
                tempString = ""
            Next

            For Each sectionToRemove As String In sectionsToRemove
                iniFile.RemoveSection(sectionToRemove)
            Next

            Dim rawINIFileContents As String = iniFile.getRawINIText

            iniFileNotes = iniFileNotes.Replace(entriesString, iniFile.Sections.Count.ToString("N0"))
            iniFile = Nothing

            Dim newINIFileContents As String = "; Version: v" & remoteINIFileVersion & vbCrLf
            newINIFileContents &= "; Last Updated On: " & Now.Date.ToLongDateString & vbCrLf
            newINIFileContents &= iniFileNotes & vbCrLf & vbCrLf
            newINIFileContents &= rawINIFileContents

            Dim streamWriter As New IO.StreamWriter(IO.Path.Combine(strLocationOfCCleaner, "winapp2.ini"))
            streamWriter.Write(newINIFileContents)
            streamWriter.Close()
            streamWriter.Dispose()

            ' These are potentially large variables, we need to trash them to free up memory.
            oldINIFileContents = Nothing
            newINIFileContents = Nothing
            rawINIFileContents = Nothing

            If Not boolSilentMode Then
                If programVariables.boolMobileMode Then
                    MsgBox("INI File Trim Complete.  A total of " & sectionsToRemove.Count.ToString("N0", Globalization.CultureInfo.CreateSpecificCulture("en-US")) & " sections were removed.", MsgBoxStyle.Information, "WinApp.ini Updater")
                Else
                    Dim msgBoxResult As MsgBoxResult = MsgBox("INI File Trim Complete.  A total of " & sectionsToRemove.Count.ToString("N0", Globalization.CultureInfo.CreateSpecificCulture("en-US")) & " sections were removed." & vbCrLf & vbCrLf & "Do you want to run CCleaner now?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "WinApp.ini Updater")

                    If msgBoxResult = Microsoft.VisualBasic.MsgBoxResult.Yes Then
                        If Environment.Is64BitOperatingSystem Then
                            Process.Start(IO.Path.Combine(strLocationOfCCleaner, "CCleaner64.exe"))
                        Else
                            Process.Start(IO.Path.Combine(strLocationOfCCleaner, "CCleaner.exe"))
                        End If
                    End If
                End If
            End If
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
            Return System.Text.RegularExpressions.Regex.Match(input, "; Version: v([0-9.A-Za-z]+)").Groups(1).Value
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
                        Else
                            Return True
                        End If
                    Else
                        Return True
                    End If
                Else
                    If Not IO.Directory.Exists(directory.Replace("%ProgramFiles%", Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles))) Then
                        If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                            sectionsToRemove.Add(iniFileSection.Name)
                            Return True
                        Else
                            Return True
                        End If
                    Else
                        Return True
                    End If
                End If
            ElseIf directory.Contains("%CommonProgramFiles%") Then
                If Environment.Is64BitOperatingSystem Then
                    If Not IO.Directory.Exists(directory.Replace("%CommonProgramFiles%", Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles))) And Not IO.Directory.Exists(directory.Replace("%CommonProgramFiles%", Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFilesX86))) Then
                        If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                            sectionsToRemove.Add(iniFileSection.Name)
                            Return True
                        Else
                            Return True
                        End If
                    Else
                        Return True
                    End If
                Else
                    If Not IO.Directory.Exists(directory.Replace("%CommonProgramFiles%", Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles))) Then
                        If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                            sectionsToRemove.Add(iniFileSection.Name)
                            Return True
                        Else
                            Return True
                        End If
                    Else
                        Return True
                    End If
                End If
            Else
                directory = translateVarsInPath(directory)
                directory = directory.Replace("*", "")

                If Not IO.Directory.Exists(directory) Then
                    If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                        sectionsToRemove.Add(iniFileSection.Name)
                        Return True
                    Else
                        Return True
                    End If
                Else
                    Return True
                End If
            End If

            Return False
        End Function

        Public Function processRegistryKey(ByVal tempString As String, ByRef sectionsToRemove As Specialized.StringCollection, ByRef iniFileSection As IniFile.IniSection) As Boolean
            Try
                If tempString.Contains(".NETFramework") Then
                    Return True
                End If

                If tempString.StartsWith("HKCU") Then
                    tempString = tempString.Replace("HKCU\", "")

                    If Environment.Is64BitOperatingSystem Then
                        If Boolean.Parse(RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry32).OpenSubKey(tempString) Is Nothing) And Boolean.Parse(RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64).OpenSubKey(tempString) Is Nothing) Then
                            If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                                sectionsToRemove.Add(iniFileSection.Name)
                                Return True
                            Else
                                Return True
                            End If
                        Else
                            Return True
                        End If
                    Else
                        If Boolean.Parse(Registry.CurrentUser.OpenSubKey(tempString) Is Nothing) Then
                            If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                                sectionsToRemove.Add(iniFileSection.Name)
                                Return True
                            Else
                                Return True
                            End If
                        Else
                            Return True
                        End If
                    End If
                ElseIf tempString.StartsWith("HKLM") Then
                    tempString = tempString.Replace("HKLM\", "")

                    If Environment.Is64BitOperatingSystem Then
                        If Boolean.Parse(RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32).OpenSubKey(tempString) Is Nothing) And Boolean.Parse(RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(tempString) Is Nothing) Then
                            If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                                sectionsToRemove.Add(iniFileSection.Name)
                                Return True
                            Else
                                Return True
                            End If
                        Else
                            Return True
                        End If
                    Else
                        If Boolean.Parse(Registry.LocalMachine.OpenSubKey(tempString) Is Nothing) Then
                            If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                                sectionsToRemove.Add(iniFileSection.Name)
                                Return True
                            Else
                                Return True
                            End If
                        Else
                            Return True
                        End If
                    End If
                ElseIf tempString.StartsWith("HKCR") Then
                    tempString = tempString.Replace("HKCR\", "")

                    If Environment.Is64BitOperatingSystem Then
                        If Boolean.Parse(RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry32).OpenSubKey(tempString) Is Nothing) And Boolean.Parse(RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry64).OpenSubKey(tempString) Is Nothing) Then
                            If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                                sectionsToRemove.Add(iniFileSection.Name)
                                Return True
                            Else
                                Return True
                            End If
                        Else
                            Return True
                        End If
                    Else
                        If Boolean.Parse(Registry.ClassesRoot.OpenSubKey(tempString) Is Nothing) Then
                            If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                                sectionsToRemove.Add(iniFileSection.Name)
                                Return True
                            Else
                                Return True
                            End If
                        Else
                            Return True
                        End If
                    End If
                ElseIf tempString.StartsWith("HKU") Then
                    tempString = tempString.Replace("HKU\", "")

                    If Environment.Is64BitOperatingSystem Then
                        If Boolean.Parse(RegistryKey.OpenBaseKey(RegistryHive.Users, RegistryView.Registry32).OpenSubKey(tempString) Is Nothing) And Boolean.Parse(RegistryKey.OpenBaseKey(RegistryHive.Users, RegistryView.Registry64).OpenSubKey(tempString) Is Nothing) Then
                            If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                                sectionsToRemove.Add(iniFileSection.Name)
                                Return True
                            Else
                                Return True
                            End If
                        Else
                            Return True
                        End If
                    Else
                        If Boolean.Parse(Registry.Users.OpenSubKey(tempString) Is Nothing) Then
                            If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                                sectionsToRemove.Add(iniFileSection.Name)
                                Return True
                            Else
                                Return True
                            End If
                        Else
                            Return True
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