Imports System.Text.RegularExpressions
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

        Private Function DoesPIDExist(PID As Integer) As Boolean
            Try
                Using searcher As New Management.ManagementObjectSearcher("root\CIMV2", String.Format("Select * FROM Win32_Process WHERE ProcessId={0}", PID))
                    Return searcher.Get.Count <> 0
                End Using
            Catch ex3 As Runtime.InteropServices.COMException
                Return False
            Catch ex As Exception
                Return False
            End Try
        End Function

        Private Sub KillProcess(PID As Integer)
            If DoesPIDExist(PID) Then Process.GetProcessById(PID).Kill()
            If DoesPIDExist(PID) Then KillProcess(PID)
        End Sub

        Private Sub SearchForProcessAndKillIt(strFileName As String, boolFullFilePathPassed As Boolean)
            Dim fullFileName As String = If(boolFullFilePathPassed, strFileName, New IO.FileInfo(strFileName).FullName)
            Dim wmiQuery As String = String.Format("Select ExecutablePath, ProcessId FROM Win32_Process WHERE ExecutablePath = '{0}'", fullFileName.AddSlashes())

            Try
                Using searcher As New Management.ManagementObjectSearcher("root\CIMV2", wmiQuery)
                    For Each queryObj As Management.ManagementObject In searcher.Get()
                        KillProcess(Integer.Parse(queryObj("ProcessId").ToString))
                    Next
                End Using
            Catch err As Exception
            End Try
        End Sub

        Private Function FindCCleaner() As String
            Try
                If Environment.Is64BitOperatingSystem Then
                    If RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey("SOFTWARE\Piriform\CCleaner", False) Is Nothing Then
                        WPFCustomMessageBox.CustomMessageBox.ShowOK("CCleaner doesn't appear to be installed on your machine.", messageBoxTitle, programConstants.strOK, Windows.MessageBoxImage.Information)
                        Return Nothing
                    Else
                        Return RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey("SOFTWARE\Piriform\CCleaner", False).GetValue(vbNullString, Nothing)
                    End If
                Else
                    If Registry.LocalMachine.OpenSubKey("SOFTWARE\Piriform\CCleaner", False) Is Nothing Then
                        WPFCustomMessageBox.CustomMessageBox.ShowOK("CCleaner doesn't appear to be installed on your machine.", messageBoxTitle, programConstants.strOK, Windows.MessageBoxImage.Information)
                        Return Nothing
                    Else
                        Return Registry.LocalMachine.OpenSubKey("SOFTWARE\Piriform\CCleaner", False).GetValue(vbNullString, Nothing)
                    End If
                End If
            Catch ex As Exception
                WPFCustomMessageBox.CustomMessageBox.ShowOK(ex.Message, messageBoxTitle, programConstants.strOK, Windows.MessageBoxImage.Information)
                Return Nothing
            End Try
        End Function

        Private Sub LoadAppSettings(ByRef boolMobileMode As Boolean, ByRef boolTrim As Boolean, ByRef boolNotifyAfterUpdateAtLogon As Boolean, ByRef boolUseSSL As Boolean, ByRef strCustomEntries As String, ByRef boolSleepOnSilentStartup As Boolean)
            If IO.File.Exists("winapp.ini updater custom entries.txt") Then
                IO.File.Move("winapp.ini updater custom entries.txt", programConstants.customEntriesFile)
            End If

            If IO.File.Exists(programConstants.configXMLFile) And IO.File.Exists(programConstants.configINIFile) Then IO.File.Delete(programConstants.configINIFile)

            If IO.File.Exists(programConstants.configXMLFile) Then
                Dim AppSettings As New AppSettings

                SyncLock programFunctions.LockObject
                    Using streamReader As New IO.StreamReader(programConstants.configXMLFile)
                        Dim xmlSerializerObject As New XmlSerializer(AppSettings.GetType)
                        AppSettings = xmlSerializerObject.Deserialize(streamReader)
                    End Using
                End SyncLock

                boolUseSSL = AppSettings.boolUseSSL
                boolSleepOnSilentStartup = AppSettings.boolSleepOnSilentStartup
                strCustomEntries = Nothing
                If Not String.IsNullOrEmpty(AppSettings.strCustomEntries) Then strCustomEntries = AppSettings.strCustomEntries.Replace(vbLf, vbCrLf)
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
                        boolUseSSL = programFunctions.GetBooleanSettingFromINIFile(iniFile, programConstants.configINIUseSSLKey)

                        iniFile.SetKeyValue(programConstants.configINISettingSection, programConstants.configINIMobileModeKey, If(boolMobileMode, 1, 0))
                        iniFile.SetKeyValue(programConstants.configINISettingSection, programConstants.configINITrimKey, If(boolTrim, 1, 0))
                        iniFile.SetKeyValue(programConstants.configINISettingSection, programConstants.configINInotifyAfterUpdateAtLogonKey, If(boolNotifyAfterUpdateAtLogon, 1, 0))
                        iniFile.SetKeyValue(programConstants.configINISettingSection, programConstants.configINIUseSSLKey, If(boolUseSSL, 1, 0))

                        iniFile.Save(programConstants.configINIFile)
                    Else
                        boolMobileMode = programFunctions.GetIntegerSettingFromINIFileAsBoolean(iniFile, programConstants.configINIMobileModeKey)
                        boolTrim = programFunctions.GetIntegerSettingFromINIFileAsBoolean(iniFile, programConstants.configINITrimKey)
                        boolNotifyAfterUpdateAtLogon = programFunctions.GetIntegerSettingFromINIFileAsBoolean(iniFile, programConstants.configINInotifyAfterUpdateAtLogonKey)
                        boolUseSSL = programFunctions.GetIntegerSettingFromINIFileAsBoolean(iniFile, programConstants.configINIUseSSLKey)
                    End If

                    Dim AppSettings As New AppSettings With {
                        .boolMobileMode = boolMobileMode,
                        .boolNotifyAfterUpdateAtLogon = boolNotifyAfterUpdateAtLogon,
                        .boolTrim = boolTrim,
                        .strCustomEntries = strCustomEntries,
                        .boolUseSSL = boolUseSSL,
                        .boolSleepOnSilentStartup = True
                    }

                    SyncLock programFunctions.LockObject
                        Using streamWriter As New IO.StreamWriter(programConstants.configXMLFile)
                            Dim xmlSerializerObject As New XmlSerializer(AppSettings.GetType)
                            xmlSerializerObject.Serialize(streamWriter, AppSettings)
                        End Using
                    End SyncLock

                    IO.File.Delete(programConstants.configINIFile)
                Else
                    Dim AppSettings As New AppSettings With {
                        .boolMobileMode = False,
                        .boolNotifyAfterUpdateAtLogon = False,
                        .boolTrim = False,
                        .strCustomEntries = "",
                        .boolUseSSL = True,
                        .boolSleepOnSilentStartup = True
                    }

                    SyncLock programFunctions.LockObject
                        Using streamWriter As New IO.StreamWriter(programConstants.configXMLFile)
                            Dim xmlSerializerObject As New XmlSerializer(AppSettings.GetType)
                            xmlSerializerObject.Serialize(streamWriter, AppSettings)
                        End Using
                    End SyncLock
                End If
            End If
        End Sub

        Private Sub MyApplication_Startup(sender As Object, e As ApplicationServices.StartupEventArgs) Handles Me.Startup
            If Environment.OSVersion.Version.Major = 5 And (Environment.OSVersion.Version.Minor = 1 Or Environment.OSVersion.Version.Minor = 2) Then
                WPFCustomMessageBox.CustomMessageBox.ShowOK("Windows XP support has been pulled from this program, this program will no longer function on Windows XP.", "YAWA2 (Yet Another WinApp2.ini) Updater", programConstants.strOK, Windows.MessageBoxImage.Error)
                e.Cancel = True
                Exit Sub
            End If

            Dim remoteINIFileVersion, localINIFileVersion As String
            Dim strLocationToSaveWinAPP2INIFile As String = Nothing
            Dim strCustomEntries As String = Nothing
            Dim boolMobileMode, boolTrim, boolNotifyAfterUpdateAtLogon, boolUseSSL, boolSleepOnSilentStartup As Boolean

            LoadAppSettings(boolMobileMode, boolTrim, boolNotifyAfterUpdateAtLogon, boolUseSSL, strCustomEntries, boolSleepOnSilentStartup)

            If My.Application.CommandLineArgs.Count = 1 Then
                Dim commandLineArgument As String = My.Application.CommandLineArgs(0).Trim

                If commandLineArgument.Equals("-silent", StringComparison.OrdinalIgnoreCase) Or commandLineArgument.Equals("/silent", StringComparison.OrdinalIgnoreCase) Then
                    If boolSleepOnSilentStartup Then Threading.Thread.Sleep(30000) ' Sleeps for thirty seconds

                    If boolMobileMode Then
                        strLocationToSaveWinAPP2INIFile = New IO.FileInfo(Windows.Forms.Application.ExecutablePath).DirectoryName
                    Else
                        strLocationToSaveWinAPP2INIFile = FindCCleaner()
                        If String.IsNullOrWhiteSpace(strLocationToSaveWinAPP2INIFile) Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    End If

                    If strLocationToSaveWinAPP2INIFile <> Nothing Then
                        If IO.File.Exists(IO.Path.Combine(strLocationToSaveWinAPP2INIFile, "winapp2.ini")) Then
                            Using streamReader As New IO.StreamReader(IO.Path.Combine(strLocationToSaveWinAPP2INIFile, "winapp2.ini"))
                                localINIFileVersion = programFunctions.GetINIVersionFromString(streamReader.ReadLine)
                            End Using
                        Else
                            localINIFileVersion = "(Not Installed)"
                        End If

                        remoteINIFileVersion = programFunctions.GetRemoteINIFileVersion()

                        If remoteINIFileVersion = programConstants.errorRetrievingRemoteINIFileVersion Then
                            WPFCustomMessageBox.CustomMessageBox.ShowOK("Error Retrieving Remote INI File Version. Please try again.", messageBoxTitle, programConstants.strOK, Windows.MessageBoxImage.Error)
                            e.Cancel = True
                            Exit Sub
                        End If

                        If IO.File.Exists(programConstants.customEntriesFile) Then
                            SyncLock programFunctions.LockObject
                                Using customEntriesFileReader As New IO.StreamReader(programConstants.customEntriesFile)
                                    strCustomEntries = customEntriesFileReader.ReadToEnd.Trim
                                End Using

                                Dim AppSettings As New AppSettings
                                Using streamReader As New IO.StreamReader(programConstants.configXMLFile)
                                    Dim xmlSerializerObject As New XmlSerializer(AppSettings.GetType)
                                    AppSettings = xmlSerializerObject.Deserialize(streamReader)
                                End Using

                                AppSettings.strCustomEntries = strCustomEntries

                                Using streamWriter As New IO.StreamWriter(programConstants.configXMLFile)
                                    Dim xmlSerializerObject As New XmlSerializer(AppSettings.GetType)
                                    xmlSerializerObject.Serialize(streamWriter, AppSettings)
                                End Using

                                IO.File.Delete(programConstants.customEntriesFile)
                            End SyncLock
                        End If

                        If remoteINIFileVersion.Trim.Equals(localINIFileVersion.Trim, StringComparison.OrdinalIgnoreCase) Then
                            e.Cancel = True
                            Exit Sub
                        Else
                            Dim remoteINIFileData As String = Nothing
                            Dim httpHelper As HttpHelper = internetFunctions.CreateNewHTTPHelperObject()

                            If httpHelper.GetWebData(programConstants.WinApp2INIFileURL, remoteINIFileData, False) Then
                                Using streamWriter As New IO.StreamWriter(IO.Path.Combine(strLocationToSaveWinAPP2INIFile, "winapp2.ini"))
                                    streamWriter.Write(If(String.IsNullOrWhiteSpace(strCustomEntries), remoteINIFileData, remoteINIFileData & vbCrLf & strCustomEntries & vbCrLf))
                                End Using
                            Else
                                WPFCustomMessageBox.CustomMessageBox.ShowOK("There was an error while downloading the WinApp2.ini file.", messageBoxTitle, programConstants.strOK, Windows.MessageBoxImage.Information)
                                e.Cancel = True
                                Exit Sub
                            End If
                        End If

                        If boolTrim Then
                            programFunctions.TrimINIFile(strLocationToSaveWinAPP2INIFile, remoteINIFileVersion, True)
                        End If

                        If boolNotifyAfterUpdateAtLogon AndAlso WPFCustomMessageBox.CustomMessageBox.ShowYesNo("The CCleaner WinApp2.ini file has been updated." & vbCrLf & vbCrLf & "Do you want to run CCleaner now?", "WinApp2.ini File Updated", programConstants.strYes, programConstants.strNo, Windows.MessageBoxImage.Question) = Windows.MessageBoxResult.Yes Then
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