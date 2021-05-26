Imports Microsoft.Win32
Imports System.Security.Principal
Imports System.Text.RegularExpressions
Imports System.Xml.Serialization

Namespace programConstants
    Module programConstants
        Public Const errorRetrievingRemoteINIFileVersion As String = "Error Retrieving Remote INI File Version"
        Public Const customEntriesFile As String = "YAWA2 Updater Custom Entries.txt"
        Public Const configINIFile As String = "YAWA2 Updater Config.ini"
        Public Const configXMLFile As String = "YAWA2 Updater Config.xml"

        Public Const WinApp2INIFileURL As String = "https://raw.githubusercontent.com/MoscaDotTo/Winapp2/master/Winapp2.ini"

        Public Const configINISettingSection As String = "Configuration"
        Public Const configINICustomEntriesKey As String = "customEntries"
        Public Const configINIMobileModeKey As String = "MobileMode"
        Public Const configINITrimKey As String = "trim"
        Public Const configINIUseSSLKey As String = "useSSL"
        Public Const configINInotifyAfterUpdateAtLogonKey As String = "notifyAfterUpdateAtLogon"
    End Module
End Namespace

Public Structure AppSettings
    Public boolMobileMode, boolTrim, boolNotifyAfterUpdateAtLogon, boolSleepOnSilentStartup As Boolean
    Public strCustomEntries As String
End Structure

Namespace programFunctions
    Module functions
        Public ReadOnly osVersionString As String = Environment.OSVersion.Version.Major & "." & Environment.OSVersion.Version.Minor
        Public ReadOnly LockObject As New Object

        Public Function CanIWriteToTheCurrentDirectory() As Boolean
            Return Check_for_Update_Stuff.CheckFolderPermissionsByACLs(New IO.FileInfo(Application.ExecutablePath).DirectoryName)
        End Function

        Public Function AreWeAnAdministrator() As Boolean
            Try
                Dim principal As New WindowsPrincipal(WindowsIdentity.GetCurrent())
                Return principal.IsInRole(WindowsBuiltInRole.Administrator)
            Catch ex As Exception
                Return False
            End Try
        End Function

        Public Function IsBase64(base64String As String) As Boolean
            If base64String.Replace(" ", "").Length Mod 4 <> 0 Then Return False
            Try
                Convert.FromBase64String(base64String)
                Return True
            Catch exception As FormatException ' Handle the exception
                Return False
            End Try
        End Function

        Public Function ConvertToBase64(input As String) As String
            Return Convert.ToBase64String(Text.Encoding.UTF8.GetBytes(input))
        End Function

        Public Function ConvertFromBase64(input As String) As String
            Return Text.Encoding.UTF8.GetString(Convert.FromBase64String(input))
        End Function

        Public Sub RemoveSettingFromINIFile(settingToRemove As String)
            ' Check if the INI file exists.
            If IO.File.Exists(programConstants.configINIFile) Then
                ' Yes, it does exist; now let's load the file into memory.

                Dim iniFile As New IniFile() ' First we create an IniFile class object.
                iniFile.LoadINIFileFromFile(programConstants.configINIFile) ' Now we load the data into memory.

                Dim temp As String = iniFile.GetKeyValue(programConstants.configINISettingSection, settingToRemove) ' Load the data from the INI object into local program memory.
                If String.IsNullOrWhiteSpace(temp) Then
                    iniFile.RemoveKey(programConstants.configINISettingSection, settingToRemove) ' Remove the setting from the INI file in memory.
                End If

                iniFile.Save(programConstants.configINIFile) ' Save the data to disk.
                iniFile = Nothing ' And destroy the INIFile object.
            End If
        End Sub

        Public Function GetBooleanSettingFromINIFile(ByRef iniFile As IniFile, strSetting As String) As Boolean
            Dim boolValue As Boolean
            If Not Boolean.TryParse(iniFile.GetKeyValue(programConstants.configINISettingSection, strSetting), boolValue) Then
                boolValue = False
            End If
            Return boolValue
        End Function

        Public Function GetIntegerSettingFromINIFileAsBoolean(ByRef iniFile As IniFile, strSetting As String) As Boolean
            Dim intValue As Integer
            If Not Integer.TryParse(iniFile.GetKeyValue(programConstants.configINISettingSection, strSetting), intValue) Then
                intValue = 0
            End If
            Return intValue = 1
        End Function

        Private Function DoesINISettingExist(ByRef iniFile As IniFile, strSetting As String, ByRef strRawValue As String) As Boolean
            strRawValue = iniFile.GetKeyValue(programConstants.configINISettingSection, strSetting)
            Return Not String.IsNullOrWhiteSpace(strRawValue)
        End Function

        Public Function GetINISettingType(ByRef iniFile As IniFile, strSetting As String) As SettingType
            Dim strRawValue As String = Nothing
            Dim boolTestValue As Boolean
            Dim intTestValue As Integer

            If DoesINISettingExist(iniFile, strSetting, strRawValue) Then
                strRawValue = strRawValue.Trim

                If Boolean.TryParse(strRawValue, boolTestValue) Then
                    Return SettingType.bool
                ElseIf Integer.TryParse(strRawValue, intTestValue) Then
                    Return SettingType.number
                Else
                    Return SettingType.unknown
                End If
            Else
                Return SettingType.null
            End If
        End Function

        Public Enum SettingType As Byte
            bool
            number
            null
            unknown
        End Enum

        Public Sub SaveSettingToAppSettingsXMLFile(AppSettingType As AppSettingType, boolValue As Boolean)
            SyncLock LockObject
                Dim AppSettings As New AppSettings
                Using streamReader As New IO.StreamReader(programConstants.configXMLFile)
                    Dim xmlSerializerObject As New XmlSerializer(AppSettings.GetType)
                    AppSettings = xmlSerializerObject.Deserialize(streamReader)
                End Using

                If AppSettingType = AppSettingType.boolMobileMode Then
                    AppSettings.boolMobileMode = boolValue
                ElseIf AppSettingType = AppSettingType.boolTrim Then
                    AppSettings.boolTrim = boolValue
                ElseIf AppSettingType = AppSettingType.boolNotifyAfterUpdateAtLogon Then
                    AppSettings.boolNotifyAfterUpdateAtLogon = boolValue
                ElseIf AppSettingType = AppSettingType.boolSleepOnSilentStartup Then
                    AppSettings.boolSleepOnSilentStartup = boolValue
                End If

                Using streamWriter As New IO.StreamWriter(programConstants.configXMLFile)
                    Dim xmlSerializerObject As New XmlSerializer(AppSettings.GetType)
                    xmlSerializerObject.Serialize(streamWriter, AppSettings)
                End Using
            End SyncLock
        End Sub

        Public Sub SaveCustomEntriesToAppSettingsXMLFile(strValue As String)
            SyncLock LockObject
                Dim AppSettings As New AppSettings
                Using streamReader As New IO.StreamReader(programConstants.configXMLFile)
                    Dim xmlSerializerObject As New XmlSerializer(AppSettings.GetType)
                    AppSettings = xmlSerializerObject.Deserialize(streamReader)
                End Using

                AppSettings.strCustomEntries = strValue

                Using streamWriter As New IO.StreamWriter(programConstants.configXMLFile)
                    Dim xmlSerializerObject As New XmlSerializer(AppSettings.GetType)
                    xmlSerializerObject.Serialize(streamWriter, AppSettings)
                End Using
            End SyncLock
        End Sub

        Public Enum AppSettingType As Byte
            boolMobileMode
            boolTrim
            boolNotifyAfterUpdateAtLogon
            boolSleepOnSilentStartup
        End Enum

        Public Function LoadSettingFromINIFile(ByVal settingKey As String, ByRef settingKeyValue As String) As Boolean
            Try
                Dim iniFile As New IniFile() ' First we create an IniFile class object.

                ' Check if the INI file exists.
                If IO.File.Exists(programConstants.configINIFile) Then
                    iniFile.LoadINIFileFromFile(programConstants.configINIFile) ' Yes, it does exist; now let's load the file into memory.

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

        Public Function GetRemoteINIFileVersion() As String
            Try
                Dim httpHelper As HttpHelper = internetFunctions.CreateNewHTTPHelperObject()
                Dim strINIFileData As String = Nothing

                If httpHelper.GetWebData(programConstants.WinApp2INIFileURL, strINIFileData, 0, 2048, False) Then
                    Return GetINIVersionFromString(strINIFileData)
                Else
                    Return programConstants.errorRetrievingRemoteINIFileVersion
                End If
            Catch ex As Exception
                Return programConstants.errorRetrievingRemoteINIFileVersion
            End Try
        End Function

        Private Function OsVersionCheck(tempString As String) As Boolean
            If Environment.OSVersion.Version.Major = 10 Then
                If tempString.Contains("6.2") Or tempString.Contains("10.0") Then Return True
            Else
                If tempString.Contains(osVersionString) Then Return True
            End If
            Return False
        End Function

        Public Sub TrimINIFile(strLocationOfCCleaner As String, remoteINIFileVersion As String, boolSilentMode As Boolean)
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
            iniFile.LoadINIFileFromFile(IO.Path.Combine(strLocationOfCCleaner, "winapp2.ini"))

            For Each iniFileSection As IniFile.IniSection In iniFile.Sections
                tempString = iniFile.GetKeyValue(iniFileSection.Name, "DetectOS")
                If Not String.IsNullOrWhiteSpace(tempString) Then
                    If Not OsVersionCheck(tempString) Then
                        If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                            sectionsToRemove.Add(iniFileSection.Name)
                            Continue For
                        End If
                    End If
                End If
                tempString = Nothing

                tempString = iniFile.GetKeyValue(iniFileSection.Name, "Detect")
                If Not String.IsNullOrWhiteSpace(tempString) Then
                    If ProcessRegistryKey(tempString, sectionsToRemove, iniFileSection) Then
                        Continue For
                    End If
                End If
                tempString = Nothing

                tempString = iniFile.GetKeyValue(iniFileSection.Name, "DetectFile")
                If Not String.IsNullOrWhiteSpace(tempString) Then
                    If ProcessFilePath(tempString, sectionsToRemove, iniFileSection) Then
                        Continue For
                    End If
                End If
                tempString = Nothing

                tempString = iniFile.GetKeyValue(iniFileSection.Name, "Detect1")
                If Not String.IsNullOrWhiteSpace(tempString) Then
                    If ProcessRegistryKey(tempString, sectionsToRemove, iniFileSection) Then
                        Continue For
                    End If
                End If
                tempString = Nothing

                tempString = iniFile.GetKeyValue(iniFileSection.Name, "DetectFile1")
                If Not String.IsNullOrWhiteSpace(tempString) Then
                    If ProcessFilePath(tempString, sectionsToRemove, iniFileSection) Then
                        Continue For
                    End If
                End If
                tempString = Nothing

                tempString = iniFile.GetKeyValue(iniFileSection.Name, "FileKey1")
                If Not String.IsNullOrWhiteSpace(tempString) Then
                    If ProcessFilePath(tempString, sectionsToRemove, iniFileSection) Then
                        Continue For
                    End If
                End If
                tempString = Nothing

                tempString = iniFile.GetKeyValue(iniFileSection.Name, "RegKey1")
                If Not String.IsNullOrWhiteSpace(tempString) Then
                    If ProcessRegistryKey(tempString, sectionsToRemove, iniFileSection) Then
                        Continue For
                    End If
                End If
                tempString = Nothing
            Next

            For Each sectionToRemove As String In sectionsToRemove
                iniFile.RemoveSection(sectionToRemove)
            Next

            Dim rawINIFileContents As String = iniFile.GetRawINIText.Trim

            iniFileNotes = iniFileNotes.Replace(entriesString, iniFile.Sections.Count.ToString("N0"))
            iniFile = Nothing

            Dim newINIFileContents As String = "; Version: " & remoteINIFileVersion & vbCrLf
            newINIFileContents &= "; Last Updated On: " & Now.Date.ToLongDateString & vbCrLf
            newINIFileContents &= iniFileNotes & DoubleCRLF
            newINIFileContents &= rawINIFileContents & vbCrLf

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

            Dim AppSettings As New AppSettings
            SyncLock LockObject
                Using streamReader As New IO.StreamReader(programConstants.configXMLFile)
                    Dim xmlSerializerObject As New XmlSerializer(AppSettings.GetType)
                    AppSettings = xmlSerializerObject.Deserialize(streamReader)
                End Using
            End SyncLock

            If Not boolSilentMode Then
                If Not AppSettings.boolMobileMode AndAlso MsgBox("INI File Trim Complete. A total of " & sectionsToRemove.Count.ToString("N0", Globalization.CultureInfo.CreateSpecificCulture("en-US")) & " sections were removed." & DoubleCRLF & "Do you want to run CCleaner now?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "WinApp.ini Updater") = MsgBoxResult.Yes Then
                    RunCCleaner(strLocationOfCCleaner)
                Else
                    MsgBox("INI File Trim Complete.  A total of " & sectionsToRemove.Count.ToString("N0", Globalization.CultureInfo.CreateSpecificCulture("en-US")) & " sections were removed.", MsgBoxStyle.Information, "WinApp.ini Updater")
                End If
            End If
        End Sub

        Public Sub RunCCleaner(strLocationOfCCleaner As String)
            Process.Start(IO.Path.Combine(strLocationOfCCleaner, If(Environment.Is64BitOperatingSystem, "CCleaner64.exe", "CCleaner.exe")))
        End Sub

        Private Function TranslateVarsInPath(input As String) As String
            input = input.Replace("%AppData%", Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData))
            input = input.Replace("%LocalAppData%", Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData))
            input = input.Replace("%Documents%", Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments))
            input = input.Replace("%CommonAppData%", Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData))
            input = input.Replace("%SystemDrive%", Environment.GetFolderPath(Environment.SpecialFolder.Windows).Replace("\Windows", ""))
            input = input.Replace("%WinDir%", Environment.GetFolderPath(Environment.SpecialFolder.Windows))
            Return input
        End Function

        Public Function GetINIVersionFromString(input As String) As String
            ' Special Regular Expression to extract the version of the INI file from the INI file's raw text.
            Return Regex.Match(input, "; Version: ([0-9.A-Za-z]+)").Groups(1).Value
        End Function

        Public Function ProcessFilePath(ByVal tempString As String, ByRef sectionsToRemove As Specialized.StringCollection, ByRef iniFileSection As IniFile.IniSection) As Boolean
            Dim directory As String = tempString.Split("|")(0).Replace("*", "")

            If directory.CaseInsensitiveContains("%ProgramFiles%") Then
                If Environment.Is64BitOperatingSystem Then
                    If Not IO.Directory.Exists(directory.CaseInsensitiveReplace("%ProgramFiles%", Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles))) And Not IO.Directory.Exists(directory.Replace("%ProgramFiles%", Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86))) Then
                        If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                            sectionsToRemove.Add(iniFileSection.Name)
                            Return True
                        Else : Return True
                        End If
                    Else : Return True
                    End If
                Else
                    If Not IO.Directory.Exists(directory.CaseInsensitiveReplace("%ProgramFiles%", Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles))) Then
                        If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                            sectionsToRemove.Add(iniFileSection.Name)
                            Return True
                        Else : Return True
                        End If
                    Else : Return True
                    End If
                End If
            ElseIf directory.CaseInsensitiveContains("%CommonProgramFiles%") Then
                If Environment.Is64BitOperatingSystem Then
                    If Not IO.Directory.Exists(directory.CaseInsensitiveReplace("%CommonProgramFiles%", Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles))) And Not IO.Directory.Exists(directory.Replace("%CommonProgramFiles%", Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFilesX86))) Then
                        If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                            sectionsToRemove.Add(iniFileSection.Name)
                            Return True
                        Else : Return True
                        End If
                    Else : Return True
                    End If
                Else
                    If Not IO.Directory.Exists(directory.CaseInsensitiveReplace("%CommonProgramFiles%", Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles))) Then
                        If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                            sectionsToRemove.Add(iniFileSection.Name)
                            Return True
                        Else : Return True
                        End If
                    Else : Return True
                    End If
                End If
            Else
                directory = TranslateVarsInPath(directory).Replace("*", "")

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

        Public Function ProcessRegistryKey(ByVal tempString As String, ByRef sectionsToRemove As Specialized.StringCollection, ByRef iniFileSection As IniFile.IniSection) As Boolean
            Try
                Dim regKey1, regKey2 As RegistryKey
                If tempString.Contains(".NETFramework") Then Return True

                If tempString.StartsWith("HKCU", StringComparison.OrdinalIgnoreCase) Then
                    tempString = tempString.CaseInsensitiveReplace("HKCU\", "")

                    If Environment.Is64BitOperatingSystem Then
                        regKey1 = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry32)
                        regKey2 = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64)

                        If Boolean.Parse(regKey1.OpenSubKey(tempString) Is Nothing) And Boolean.Parse(regKey2.OpenSubKey(tempString) Is Nothing) Then
                            If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                                sectionsToRemove.Add(iniFileSection.Name)
                                regKey1.Dispose()
                                regKey2.Dispose()
                                Return True
                            Else
                                regKey1.Dispose()
                                regKey2.Dispose()
                                Return True
                            End If
                        Else
                            regKey1.Dispose()
                            regKey2.Dispose()
                            Return True
                        End If
                    Else
                        regKey1 = Registry.CurrentUser

                        If Boolean.Parse(regKey1.OpenSubKey(tempString) Is Nothing) Then
                            If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                                sectionsToRemove.Add(iniFileSection.Name)
                                regKey1.Dispose()
                                Return True
                            Else
                                regKey1.Dispose()
                                Return True
                            End If
                        Else
                            regKey1.Dispose()
                            Return True
                        End If
                    End If
                ElseIf tempString.StartsWith("HKLM", StringComparison.OrdinalIgnoreCase) Then
                    tempString = tempString.CaseInsensitiveReplace("HKLM\", "")

                    If Environment.Is64BitOperatingSystem Then
                        regKey1 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32)
                        regKey2 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64)

                        If Boolean.Parse(regKey1.OpenSubKey(tempString) Is Nothing) And Boolean.Parse(regKey2.OpenSubKey(tempString) Is Nothing) Then
                            If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                                sectionsToRemove.Add(iniFileSection.Name)
                                regKey1.Dispose()
                                regKey2.Dispose()
                                Return True
                            Else
                                regKey1.Dispose()
                                regKey2.Dispose()
                                Return True
                            End If
                        Else
                            regKey1.Dispose()
                            regKey2.Dispose()
                            Return True
                        End If
                    Else
                        regKey1 = Registry.LocalMachine

                        If Boolean.Parse(regKey1.OpenSubKey(tempString) Is Nothing) Then
                            If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                                sectionsToRemove.Add(iniFileSection.Name)
                                regKey1.Dispose()
                                Return True
                            Else
                                regKey1.Dispose()
                                Return True
                            End If
                        Else
                            regKey1.Dispose()
                            Return True
                        End If
                    End If
                ElseIf tempString.StartsWith("HKCR", StringComparison.OrdinalIgnoreCase) Then
                    tempString = tempString.CaseInsensitiveReplace("HKCR\", "")

                    If Environment.Is64BitOperatingSystem Then
                        regKey1 = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry32)
                        regKey2 = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry64)

                        If Boolean.Parse(regKey1.OpenSubKey(tempString) Is Nothing) And Boolean.Parse(regKey2.OpenSubKey(tempString) Is Nothing) Then
                            If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                                sectionsToRemove.Add(iniFileSection.Name)
                                regKey1.Dispose()
                                regKey2.Dispose()
                                Return True
                            Else
                                regKey1.Dispose()
                                regKey2.Dispose()
                                Return True
                            End If
                        Else
                            regKey1.Dispose()
                            regKey2.Dispose()
                            Return True
                        End If
                    Else
                        regKey1 = Registry.ClassesRoot

                        If Boolean.Parse(regKey1.OpenSubKey(tempString) Is Nothing) Then
                            If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                                sectionsToRemove.Add(iniFileSection.Name)
                                regKey1.Dispose()
                                Return True
                            Else
                                regKey1.Dispose()
                                Return True
                            End If
                        Else
                            regKey1.Dispose()
                            Return True
                        End If
                    End If
                ElseIf tempString.StartsWith("HKU", StringComparison.OrdinalIgnoreCase) Then
                    tempString = tempString.CaseInsensitiveReplace("HKU\", "")

                    If Environment.Is64BitOperatingSystem Then
                        regKey1 = RegistryKey.OpenBaseKey(RegistryHive.Users, RegistryView.Registry32)
                        regKey2 = RegistryKey.OpenBaseKey(RegistryHive.Users, RegistryView.Registry64)

                        If Boolean.Parse(regKey1.OpenSubKey(tempString) Is Nothing) And Boolean.Parse(regKey2.OpenSubKey(tempString) Is Nothing) Then
                            If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                                sectionsToRemove.Add(iniFileSection.Name)
                                regKey1.Dispose()
                                regKey2.Dispose()
                                Return True
                            Else
                                regKey1.Dispose()
                                regKey2.Dispose()
                                Return True
                            End If
                        Else
                            regKey1.Dispose()
                            regKey2.Dispose()
                            Return True
                        End If
                    Else
                        regKey1 = Registry.Users

                        If Boolean.Parse(regKey1.OpenSubKey(tempString) Is Nothing) Then
                            If Not sectionsToRemove.Contains(iniFileSection.Name) Then
                                sectionsToRemove.Add(iniFileSection.Name)
                                regKey1.Dispose()
                                Return True
                            Else
                                regKey1.Dispose()
                                Return True
                            End If
                        Else
                            regKey1.Dispose()
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