Imports System.Text.RegularExpressions
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

        Private Sub MyApplication_Startup(sender As Object, e As ApplicationServices.StartupEventArgs) Handles Me.Startup
            If Environment.OSVersion.Version.Major = 5 And (Environment.OSVersion.Version.Minor = 1 Or Environment.OSVersion.Version.Minor = 2) Then
                WPFCustomMessageBox.CustomMessageBox.ShowOK("Windows XP support has been pulled from this program, this program will no longer function on Windows XP.", "YAWA2 (Yet Another WinApp2.ini) Updater", programConstants.strOK, Windows.MessageBoxImage.Error)
                e.Cancel = True
                Exit Sub
            End If

            Dim remoteINIFileVersion, localINIFileVersion As String
            Dim strLocationToSaveWinAPP2INIFile As String = Nothing
            Dim stringCustomEntries As String = Nothing

            If IO.File.Exists(programConstants.configINIFile) Then
                Dim iniFile As New IniFile()
                iniFile.LoadINIFileFromFile(programConstants.configINIFile)

                If programFunctions.GetINISettingType(iniFile, programConstants.configINIUseSSLKey) = programFunctions.SettingType.bool Then
                    programVariables.boolMobileMode = programFunctions.GetBooleanSettingFromINIFile(iniFile, programConstants.configINIMobileModeKey)
                    programVariables.boolTrim = programFunctions.GetBooleanSettingFromINIFile(iniFile, programConstants.configINITrimKey)
                    programVariables.boolNotifyAfterUpdateAtLogon = programFunctions.GetBooleanSettingFromINIFile(iniFile, programConstants.configINInotifyAfterUpdateAtLogonKey)
                    programVariables.boolUseSSL = programFunctions.GetBooleanSettingFromINIFile(iniFile, programConstants.configINIUseSSLKey)

                    iniFile.SetKeyValue(programConstants.configINISettingSection, programConstants.configINIMobileModeKey, If(programVariables.boolMobileMode, 1, 0))
                    iniFile.SetKeyValue(programConstants.configINISettingSection, programConstants.configINITrimKey, If(programVariables.boolTrim, 1, 0))
                    iniFile.SetKeyValue(programConstants.configINISettingSection, programConstants.configINInotifyAfterUpdateAtLogonKey, If(programVariables.boolNotifyAfterUpdateAtLogon, 1, 0))
                    iniFile.SetKeyValue(programConstants.configINISettingSection, programConstants.configINIUseSSLKey, If(programVariables.boolUseSSL, 1, 0))

                    iniFile.Save(programConstants.configINIFile)
                Else
                    programVariables.boolMobileMode = programFunctions.GetIntegerSettingFromINIFileAsBoolean(iniFile, programConstants.configINIMobileModeKey)
                    programVariables.boolTrim = programFunctions.GetIntegerSettingFromINIFileAsBoolean(iniFile, programConstants.configINITrimKey)
                    programVariables.boolNotifyAfterUpdateAtLogon = programFunctions.GetIntegerSettingFromINIFileAsBoolean(iniFile, programConstants.configINInotifyAfterUpdateAtLogonKey)
                    programVariables.boolUseSSL = programFunctions.GetIntegerSettingFromINIFileAsBoolean(iniFile, programConstants.configINIUseSSLKey)
                End If
            Else
                Dim iniFile As New IniFile()
                iniFile.AddSection(programConstants.configINISettingSection)

                iniFile.SetKeyValue(programConstants.configINISettingSection, programConstants.configINIMobileModeKey, 0)
                iniFile.SetKeyValue(programConstants.configINISettingSection, programConstants.configINITrimKey, 0)
                iniFile.SetKeyValue(programConstants.configINISettingSection, programConstants.configINInotifyAfterUpdateAtLogonKey, 0)
                iniFile.SetKeyValue(programConstants.configINISettingSection, programConstants.configINIUseSSLKey, 1)

                programVariables.boolMobileMode = False
                programVariables.boolTrim = False
                programVariables.boolNotifyAfterUpdateAtLogon = False
                programVariables.boolUseSSL = True

                iniFile.Save(programConstants.configINIFile)
            End If

            If My.Application.CommandLineArgs.Count = 1 Then
                Dim commandLineArgument As String = My.Application.CommandLineArgs(0).Trim

                If commandLineArgument.Equals("-silent", StringComparison.OrdinalIgnoreCase) Or commandLineArgument.Equals("/silent", StringComparison.OrdinalIgnoreCase) Then
                    Threading.Thread.Sleep(30000) ' Sleeps for thirty seconds

                    If programVariables.boolMobileMode Then
                        strLocationToSaveWinAPP2INIFile = New IO.FileInfo(Windows.Forms.Application.ExecutablePath).DirectoryName
                    Else
                        Try
                            If Environment.Is64BitOperatingSystem Then
                                If RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey("SOFTWARE\Piriform\CCleaner", False) Is Nothing Then
                                    WPFCustomMessageBox.CustomMessageBox.ShowOK("CCleaner doesn't appear to be installed on your machine.", messageBoxTitle, programConstants.strOK, Windows.MessageBoxImage.Information)
                                    e.Cancel = True
                                    Exit Sub
                                Else
                                    strLocationToSaveWinAPP2INIFile = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey("SOFTWARE\Piriform\CCleaner", False).GetValue(vbNullString, Nothing)
                                End If
                            Else
                                If Registry.LocalMachine.OpenSubKey("SOFTWARE\Piriform\CCleaner", False) Is Nothing Then
                                    WPFCustomMessageBox.CustomMessageBox.ShowOK("CCleaner doesn't appear to be installed on your machine.", messageBoxTitle, programConstants.strOK, Windows.MessageBoxImage.Information)
                                    e.Cancel = True
                                    Exit Sub
                                Else
                                    strLocationToSaveWinAPP2INIFile = Registry.LocalMachine.OpenSubKey("SOFTWARE\Piriform\CCleaner", False).GetValue(vbNullString, Nothing)
                                End If
                            End If
                        Catch ex As Exception
                            WPFCustomMessageBox.CustomMessageBox.ShowOK(ex.Message, messageBoxTitle, programConstants.strOK, Windows.MessageBoxImage.Information)
                        End Try
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
                            Using customEntriesFileReader As New IO.StreamReader(programConstants.customEntriesFile)
                                stringCustomEntries = customEntriesFileReader.ReadToEnd.Trim
                            End Using

                            programFunctions.SaveSettingToINIFile(programConstants.configINICustomEntriesKey, programFunctions.ConvertToBase64(stringCustomEntries))
                            IO.File.Delete(programConstants.customEntriesFile)
                        Else
                            Dim iniFile As New IniFile()
                            iniFile.LoadINIFileFromFile(programConstants.configINIFile)
                            stringCustomEntries = iniFile.GetKeyValue(programConstants.configINISettingSection, programConstants.configINICustomEntriesKey)

                            If Not String.IsNullOrWhiteSpace(stringCustomEntries) Then
                                If programFunctions.IsBase64(stringCustomEntries) Then
                                    stringCustomEntries = programFunctions.ConvertFromBase64(stringCustomEntries)
                                Else
                                    stringCustomEntries = Nothing
                                End If
                            End If
                        End If

                        If remoteINIFileVersion.Trim.Equals(localINIFileVersion.Trim, StringComparison.OrdinalIgnoreCase) Then
                            e.Cancel = True
                            Exit Sub
                        Else
                            Dim remoteINIFileData As String = Nothing
                            Dim httpHelper As HttpHelper = internetFunctions.CreateNewHTTPHelperObject()

                            If httpHelper.GetWebData(programConstants.WinApp2INIFileURL, remoteINIFileData, False) Then
                                Using streamWriter As New IO.StreamWriter(IO.Path.Combine(strLocationToSaveWinAPP2INIFile, "winapp2.ini"))
                                    streamWriter.Write(If(String.IsNullOrWhiteSpace(stringCustomEntries), remoteINIFileData, remoteINIFileData & vbCrLf & stringCustomEntries & vbCrLf))
                                End Using
                            Else
                                WPFCustomMessageBox.CustomMessageBox.ShowOK("There was an error while downloading the WinApp2.ini file.", messageBoxTitle, programConstants.strOK, Windows.MessageBoxImage.Information)
                                e.Cancel = True
                                Exit Sub
                            End If
                        End If

                        If programVariables.boolTrim Then
                            programFunctions.TrimINIFile(strLocationToSaveWinAPP2INIFile, remoteINIFileVersion, True)
                        End If

                        If programVariables.boolNotifyAfterUpdateAtLogon AndAlso WPFCustomMessageBox.CustomMessageBox.ShowYesNo("The CCleaner WinApp2.ini file has been updated." & vbCrLf & vbCrLf & "Do you want to run CCleaner now?", "WinApp2.ini File Updated", programConstants.strYes, programConstants.strNo, Windows.MessageBoxImage.Question) = Windows.MessageBoxResult.Yes Then
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