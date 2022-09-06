Imports System.Xml.Serialization
Imports Microsoft.Win32

Namespace My

    ' The following events are available for MyApplication:
    ' 
    ' Startup: Raised when the application starts, before the startup form is created.
    ' Shutdown: Raised after all application forms are closed.  This event is not raised if the application terminates abnormally.
    ' UnhandledException: Raised if the application encounters an unhandled exception.
    ' StartupNextInstance: Raised when launching a single-instance application and the application is already active. 
    ' NetworkAvailabilityChanged: Raised when the network connection is connected or disconnected.
    Partial Friend Class MyApplication
        Private Const messageBoxTitle As String = "YAWA2 Updater"

        Private Function FindCCleaner() As String
            Try
                If Environment.Is64BitOperatingSystem Then
                    If RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey("SOFTWARE\Piriform\CCleaner", False) Is Nothing Then
                        MsgBox("CCleaner doesn't appear to be installed on your machine.", MsgBoxStyle.Information, messageBoxTitle)
                        Return Nothing
                    Else
                        Return RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey("SOFTWARE\Piriform\CCleaner", False).GetValue(vbNullString, Nothing)
                    End If
                Else
                    If Registry.LocalMachine.OpenSubKey("SOFTWARE\Piriform\CCleaner", False) Is Nothing Then
                        MsgBox("CCleaner doesn't appear to be installed on your machine.", MsgBoxStyle.Information, messageBoxTitle)
                        Return Nothing
                    Else
                        Return Registry.LocalMachine.OpenSubKey("SOFTWARE\Piriform\CCleaner", False).GetValue(vbNullString, Nothing)
                    End If
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Information, messageBoxTitle)
                Return Nothing
            End Try
        End Function

        Private Sub LoadAppSettings(ByRef boolMobileMode As Boolean, ByRef boolTrim As Boolean, ByRef boolNotifyAfterUpdateAtLogon As Boolean, ByRef strCustomEntries As String, ByRef boolSleepOnSilentStartup As Boolean, ByRef shortSleepOnSilentStartup As Short)
            SyncLock programFunctions.LockObject
startAgain:
                Try
                    If IO.File.Exists("winapp.ini updater custom entries.txt") Then
                        IO.File.Move("winapp.ini updater custom entries.txt", programConstants.customEntriesFile)
                    End If

                    If IO.File.Exists(programConstants.configXMLFile) And IO.File.Exists(programConstants.configINIFile) Then IO.File.Delete(programConstants.configINIFile)

                    If IO.File.Exists(programConstants.configXMLFile) Then
                        AppSettingsObject = New AppSettings

                        Try
                            AppSettingsObject = LoadSettingsFromXMLFileAppSettings()
                        Catch ex As Exception
                            IO.File.Delete(programConstants.configXMLFile)
                            GoTo startAgain
                        End Try

                        boolSleepOnSilentStartup = AppSettingsObject.boolSleepOnSilentStartup
                        shortSleepOnSilentStartup = AppSettingsObject.shortSleepOnSilentStartup
                        strCustomEntries = Nothing
                        If Not String.IsNullOrEmpty(AppSettingsObject.strCustomEntries) Then strCustomEntries = AppSettingsObject.strCustomEntries.Replace(vbLf, vbCrLf)
                    Else
                        If IO.File.Exists(programConstants.configINIFile) Then
                            Dim iniFile As New IniFile()
                            iniFile.LoadINIFileFromFile(programConstants.configINIFile)

                            strCustomEntries = iniFile.GetKeyValue(programConstants.configINISettingSection, programConstants.configINICustomEntriesKey)

                            If Not String.IsNullOrWhiteSpace(strCustomEntries) Then
                                If programFunctions.IsBase64(strCustomEntries) Then
                                    strCustomEntries = programFunctions.ConvertFromBase64(strCustomEntries)
                                Else
                                    strCustomEntries = Nothing
                                End If
                            End If

                            If programFunctions.GetINISettingType(iniFile, programConstants.configINIUseSSLKey) = programFunctions.SettingType.bool Then
                                boolMobileMode = programFunctions.GetBooleanSettingFromINIFile(iniFile, programConstants.configINIMobileModeKey)
                                boolTrim = programFunctions.GetBooleanSettingFromINIFile(iniFile, programConstants.configINITrimKey)
                                boolNotifyAfterUpdateAtLogon = programFunctions.GetBooleanSettingFromINIFile(iniFile, programConstants.configINInotifyAfterUpdateAtLogonKey)

                                iniFile.SetKeyValue(programConstants.configINISettingSection, programConstants.configINIMobileModeKey, If(boolMobileMode, 1, 0))
                                iniFile.SetKeyValue(programConstants.configINISettingSection, programConstants.configINITrimKey, If(boolTrim, 1, 0))
                                iniFile.SetKeyValue(programConstants.configINISettingSection, programConstants.configINInotifyAfterUpdateAtLogonKey, If(boolNotifyAfterUpdateAtLogon, 1, 0))

                                iniFile.Save(programConstants.configINIFile)
                            Else
                                boolMobileMode = programFunctions.GetIntegerSettingFromINIFileAsBoolean(iniFile, programConstants.configINIMobileModeKey)
                                boolTrim = programFunctions.GetIntegerSettingFromINIFileAsBoolean(iniFile, programConstants.configINITrimKey)
                                boolNotifyAfterUpdateAtLogon = programFunctions.GetIntegerSettingFromINIFileAsBoolean(iniFile, programConstants.configINInotifyAfterUpdateAtLogonKey)
                            End If

                            AppSettingsObject = New AppSettings With {
                                .boolMobileMode = boolMobileMode,
                                .boolNotifyAfterUpdateAtLogon = boolNotifyAfterUpdateAtLogon,
                                .boolTrim = boolTrim,
                                .strCustomEntries = strCustomEntries,
                                .boolSleepOnSilentStartup = True
                            }

                            SaveSettingsToXMLFile()

                            IO.File.Delete(programConstants.configINIFile)
                        Else
                            AppSettingsObject = New AppSettings With {
                                .boolMobileMode = False,
                                .boolNotifyAfterUpdateAtLogon = False,
                                .boolTrim = False,
                                .strCustomEntries = "",
                                .boolSleepOnSilentStartup = True,
                                .shortSleepOnSilentStartup = 60
                            }

                            SaveSettingsToXMLFile()
                        End If
                    End If
                Catch ex As UnauthorizedAccessException
                    Dim strFullFilePathToConfigXMLFile As String = New IO.FileInfo(programConstants.configXMLFile).FullName
                    If strFullFilePathToConfigXMLFile.CaseInsensitiveContains("onedrive") Then
                        MsgBox("An error occurred while attempting to access the application configuration settings file (YAWA2 Updater Config.xml)." & vbCrLf & vbCrLf & "This file exist in your Microsoft OneDrive, please right-click on the file and click on ""Always keep on this device"".", MsgBoxStyle.Critical, "YAWA2 (Yet Another WinApp2.ini) Updater")
                        SelectFileInWindowsExplorer(strFullFilePathToConfigXMLFile)
                    End If
                End Try
            End SyncLock
        End Sub

        Private Sub SelectFileInWindowsExplorer(strFullPath As String)
            If Not String.IsNullOrEmpty(strFullPath) AndAlso IO.File.Exists(strFullPath) Then
                Dim pidlList As IntPtr = NativeMethod.NativeMethods.ILCreateFromPathW(strFullPath)

                If Not pidlList.Equals(IntPtr.Zero) Then
                    Try
                        NativeMethod.NativeMethods.SHOpenFolderAndSelectItems(pidlList, 0, IntPtr.Zero, 0)
                    Finally
                        NativeMethod.NativeMethods.ILFree(pidlList)
                    End Try
                End If
            End If
        End Sub

        Private Sub MyApplication_Startup(sender As Object, e As ApplicationServices.StartupEventArgs) Handles Me.Startup
            If Environment.OSVersion.Version.Major = 5 And (Environment.OSVersion.Version.Minor = 1 Or Environment.OSVersion.Version.Minor = 2) Then
                MsgBox("Windows XP support has been pulled from this program, this program will no longer function on Windows XP.", MsgBoxStyle.Critical, "YAWA2 (Yet Another WinApp2.ini) Updater")
                e.Cancel = True
                Exit Sub
            End If

            Dim remoteINIFileVersion, localINIFileVersion As String
            Dim strLocationToSaveWinAPP2INIFile As String = Nothing
            Dim strCustomEntries As String = Nothing
            Dim boolMobileMode, boolTrim, boolNotifyAfterUpdateAtLogon, boolSleepOnSilentStartup As Boolean
            Dim shortSleepOnSilentStartup As Short

            LoadAppSettings(boolMobileMode, boolTrim, boolNotifyAfterUpdateAtLogon, strCustomEntries, boolSleepOnSilentStartup, shortSleepOnSilentStartup)

            If shortSleepOnSilentStartup = 0 Then
                shortSleepOnSilentStartup = 60
                AppSettingsObject.shortSleepOnSilentStartup = 60
                SaveSettingsToXMLFile()
            End If

            If Application.CommandLineArgs.Count = 1 Then
                Dim commandLineArgument As String = Application.CommandLineArgs(0).Trim

                If commandLineArgument.Equals("-silent", StringComparison.OrdinalIgnoreCase) Or commandLineArgument.Equals("/silent", StringComparison.OrdinalIgnoreCase) Then
                    If boolSleepOnSilentStartup Then Threading.Thread.Sleep(shortSleepOnSilentStartup * 1000)

                    If boolMobileMode Then
                        strLocationToSaveWinAPP2INIFile = New IO.FileInfo(Windows.Forms.Application.ExecutablePath).DirectoryName
                    Else
                        strLocationToSaveWinAPP2INIFile = FindCCleaner()
                        If String.IsNullOrWhiteSpace(strLocationToSaveWinAPP2INIFile) Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    End If

                    If Not String.IsNullOrWhiteSpace(strLocationToSaveWinAPP2INIFile) Then
                        If IO.File.Exists(IO.Path.Combine(strLocationToSaveWinAPP2INIFile, "winapp2.ini")) Then
                            Using streamReader As New IO.StreamReader(IO.Path.Combine(strLocationToSaveWinAPP2INIFile, "winapp2.ini"))
                                localINIFileVersion = programFunctions.GetINIVersionFromString(streamReader.ReadLine)
                            End Using
                        Else
                            localINIFileVersion = "(Not Installed)"
                        End If

                        remoteINIFileVersion = programFunctions.GetRemoteINIFileVersion()

                        If remoteINIFileVersion = programConstants.errorRetrievingRemoteINIFileVersion Then
                            MsgBox("Error Retrieving Remote INI File Version. Please try again.", MsgBoxStyle.Critical, messageBoxTitle)
                            e.Cancel = True
                            Exit Sub
                        End If

                        If IO.File.Exists(programConstants.customEntriesFile) Then
                            Using customEntriesFileReader As New IO.StreamReader(programConstants.customEntriesFile)
                                strCustomEntries = customEntriesFileReader.ReadToEnd.Trim
                            End Using

                            LoadSettingsFromXMLFileAppSettings()
                            AppSettingsObject.strCustomEntries = strCustomEntries
                            SaveSettingsToXMLFile()

                            IO.File.Delete(programConstants.customEntriesFile)
                        End If

                        If remoteINIFileVersion.Trim.Equals(localINIFileVersion.Trim, StringComparison.OrdinalIgnoreCase) Then
                            e.Cancel = True
                            Exit Sub
                        Else
                            Dim remoteINIFileData As String = Nothing
                            Dim httpHelper As HttpHelper = internetFunctions.CreateNewHTTPHelperObject()

                            If httpHelper.GetWebData(programConstants.WinApp2INIFileURL, remoteINIFileData, False) Then
                                Using streamWriter As New IO.StreamWriter(IO.Path.Combine(strLocationToSaveWinAPP2INIFile, "winapp2.ini"))
                                    If String.IsNullOrWhiteSpace(strCustomEntries) Then
                                        streamWriter.Write(remoteINIFileData.Trim & vbCrLf)
                                    Else
                                        streamWriter.Write(remoteINIFileData.Trim & DoubleCRLF & strCustomEntries & vbCrLf)
                                    End If
                                End Using
                            Else
                                MsgBox("There was an error while downloading the WinApp2.ini file.", MsgBoxStyle.Information, messageBoxTitle)
                                e.Cancel = True
                                Exit Sub
                            End If
                        End If

                        If boolTrim Then
                            programFunctions.TrimINIFile(strLocationToSaveWinAPP2INIFile, remoteINIFileVersion, True)
                        End If

                        If boolNotifyAfterUpdateAtLogon AndAlso MsgBox("The CCleaner WinApp2.ini file has been updated." & DoubleCRLF & "Do you want to run CCleaner now?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "WinApp2.ini File Updated") = MsgBoxResult.Yes Then
                            Process.Start(IO.Path.Combine(strLocationToSaveWinAPP2INIFile, If(Environment.Is64BitOperatingSystem, "CCleaner64.exe", "CCleaner.exe")))
                        End If
                    End If
                ElseIf commandLineArgument.Equals("-update", StringComparison.OrdinalIgnoreCase) Then
                    DoUpdateAtStartup()
                End If

                e.Cancel = True
                Exit Sub
            End If
        End Sub
    End Class
End Namespace