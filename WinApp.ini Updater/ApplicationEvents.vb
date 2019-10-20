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

        Private Function doesPIDExist(PID As Integer) As Boolean
            Try
                Using searcher As New Management.ManagementObjectSearcher("root\CIMV2", String.Format("Select * FROM Win32_Process WHERE ProcessId={0}", PID))
                    Return If(searcher.Get.Count = 0, False, True)
                End Using
            Catch ex3 As Runtime.InteropServices.COMException
                Return False
            Catch ex As Exception
                Return False
            End Try
        End Function

        Private Sub killProcess(PID As Integer)
            If doesPIDExist(PID) Then Process.GetProcessById(PID).Kill()
            If doesPIDExist(PID) Then killProcess(PID)
        End Sub

        Private Sub searchForProcessAndKillIt(strFileName As String, boolFullFilePathPassed As Boolean)
            Dim fullFileName As String = If(boolFullFilePathPassed, strFileName, New IO.FileInfo(strFileName).FullName)
            Dim wmiQuery As String = String.Format("Select ExecutablePath, ProcessId FROM Win32_Process WHERE ExecutablePath = '{0}'", fullFileName.addSlashes())

            Try
                Using searcher As New Management.ManagementObjectSearcher("root\CIMV2", wmiQuery)
                    For Each queryObj As Management.ManagementObject In searcher.Get()
                        killProcess(Integer.Parse(queryObj("ProcessId").ToString))
                    Next
                End Using
            Catch ex3 As Runtime.InteropServices.COMException
            Catch err As Management.ManagementException
                ' Does nothing
            End Try
        End Sub

        Private Sub MyApplication_Startup(sender As Object, e As ApplicationServices.StartupEventArgs) Handles Me.Startup
            If Environment.OSVersion.Version.Major = 5 And (Environment.OSVersion.Version.Minor = 1 Or Environment.OSVersion.Version.Minor = 2) Then
                MsgBox("Windows XP support has been pulled from this program, this program will no longer function on Windows XP.", MsgBoxStyle.Critical, "YAWA2 (Yet Another WinApp2.ini) Updater")
                e.Cancel = True
                Exit Sub
            End If

            Dim remoteINIFileVersion, localINIFileVersion As String
            Dim strLocationToSaveWinAPP2INIFile As String = Nothing
            Dim stringCustomEntries As String = Nothing

            If IO.File.Exists(programConstants.configINIFile) Then
                Dim iniFile As New IniFile()
                iniFile.loadINIFileFromFile(programConstants.configINIFile)

                If Not Boolean.TryParse(iniFile.GetKeyValue(programConstants.configINISettingSection, programConstants.configINIMobileModeKey), programVariables.boolMobileMode) Then
                    programVariables.boolMobileMode = False
                End If

                If Boolean.TryParse(iniFile.GetKeyValue(programConstants.configINISettingSection, programConstants.configINITrimKey), programVariables.boolTrim) = False Then
                    programVariables.boolTrim = False
                End If

                If Not Boolean.TryParse(iniFile.GetKeyValue(programConstants.configINISettingSection, programConstants.configINInotifyAfterUpdateAtLogonKey), programVariables.boolNotifyAfterUpdateAtLogon) Then
                    programVariables.boolNotifyAfterUpdateAtLogon = False
                End If
            Else
                Dim iniFile As New IniFile()
                iniFile.AddSection(programConstants.configINISettingSection)

                iniFile.SetKeyValue(programConstants.configINISettingSection, programConstants.configINIMobileModeKey, False)
                iniFile.SetKeyValue(programConstants.configINISettingSection, programConstants.configINITrimKey, False)
                iniFile.SetKeyValue(programConstants.configINISettingSection, programConstants.configINInotifyAfterUpdateAtLogonKey, False)

                programVariables.boolMobileMode = False
                programVariables.boolTrim = False
                programVariables.boolNotifyAfterUpdateAtLogon = False

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
                                    MsgBox("CCleaner doesn't appear to be installed on your machine.", MsgBoxStyle.Information, messageBoxTitle)
                                    e.Cancel = True
                                    Exit Sub
                                Else
                                    strLocationToSaveWinAPP2INIFile = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey("SOFTWARE\Piriform\CCleaner", False).GetValue(vbNullString, Nothing)
                                End If
                            Else
                                If Registry.LocalMachine.OpenSubKey("SOFTWARE\Piriform\CCleaner", False) Is Nothing Then
                                    MsgBox("CCleaner doesn't appear to be installed on your machine.", MsgBoxStyle.Information, messageBoxTitle)
                                    e.Cancel = True
                                    Exit Sub
                                Else
                                    strLocationToSaveWinAPP2INIFile = Registry.LocalMachine.OpenSubKey("SOFTWARE\Piriform\CCleaner", False).GetValue(vbNullString, Nothing)
                                End If
                            End If
                        Catch ex As Exception
                            MsgBox(ex.Message)
                        End Try
                    End If

                    If strLocationToSaveWinAPP2INIFile <> Nothing Then
                        If IO.File.Exists(IO.Path.Combine(strLocationToSaveWinAPP2INIFile, "winapp2.ini")) Then
                            Using streamReader As New IO.StreamReader(IO.Path.Combine(strLocationToSaveWinAPP2INIFile, "winapp2.ini"))
                                localINIFileVersion = programFunctions.getINIVersionFromString(streamReader.ReadLine)
                            End Using
                        Else
                            localINIFileVersion = "(Not Installed)"
                        End If

                        remoteINIFileVersion = programFunctions.getRemoteINIFileVersion()

                        If remoteINIFileVersion = programConstants.errorRetrievingRemoteINIFileVersion Then
                            MsgBox("Error Retrieving Remote INI File Version.  Please try again.", MsgBoxStyle.Critical, messageBoxTitle)
                            e.Cancel = True
                            Exit Sub
                        End If

                        If IO.File.Exists(programConstants.customEntriesFile) Then
                            Using customEntriesFileReader As New IO.StreamReader(programConstants.customEntriesFile)
                                stringCustomEntries = customEntriesFileReader.ReadToEnd.Trim
                            End Using

                            programFunctions.saveSettingToINIFile(programConstants.configINICustomEntriesKey, programFunctions.convertToBase64(stringCustomEntries))
                            IO.File.Delete(programConstants.customEntriesFile)
                        Else
                            Dim iniFile As New IniFile()
                            iniFile.loadINIFileFromFile(programConstants.configINIFile)
                            stringCustomEntries = iniFile.GetKeyValue(programConstants.configINISettingSection, programConstants.configINICustomEntriesKey)

                            If Not String.IsNullOrWhiteSpace(stringCustomEntries) Then
                                If programFunctions.isBase64(stringCustomEntries) Then
                                    stringCustomEntries = programFunctions.convertFromBase64(stringCustomEntries)
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
                            Dim httpHelper As httpHelper = internetFunctions.createNewHTTPHelperObject()

                            If httpHelper.getWebData(programConstants.WinApp2INIFileURL, remoteINIFileData, False) Then
                                Using streamWriter As New IO.StreamWriter(IO.Path.Combine(strLocationToSaveWinAPP2INIFile, "winapp2.ini"))
                                    streamWriter.Write(If(String.IsNullOrWhiteSpace(stringCustomEntries), remoteINIFileData, remoteINIFileData & vbCrLf & stringCustomEntries & vbCrLf))
                                End Using
                            Else
                                MsgBox("There was an error while downloading the WinApp2.ini file.", MsgBoxStyle.Information, "YAWA2 Updater")
                                e.Cancel = True
                                Exit Sub
                            End If
                        End If

                        If programVariables.boolTrim Then
                            programFunctions.trimINIFile(strLocationToSaveWinAPP2INIFile, remoteINIFileVersion, True)
                        End If

                        If programVariables.boolNotifyAfterUpdateAtLogon Then
                            Dim msgBoxResult As MsgBoxResult = MsgBox("The CCleaner WinApp2.ini file has been updated." & vbCrLf & vbCrLf & "New Remote INI File Version: " & remoteINIFileVersion & vbCrLf & vbCrLf & "Do you want to run CCleaner now?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "WinApp2.ini File Updated")

                            If msgBoxResult = Microsoft.VisualBasic.MsgBoxResult.Yes Then
                                Process.Start(IO.Path.Combine(strLocationToSaveWinAPP2INIFile, If(Environment.Is64BitOperatingSystem, "CCleaner64.exe", "CCleaner.exe")))
                            End If
                        End If
                    End If
                ElseIf commandLineArgument.Equals("-update", StringComparison.OrdinalIgnoreCase) Then
                    Dim currentProcessFileName As String = New IO.FileInfo(Windows.Forms.Application.ExecutablePath).Name

                    If currentProcessFileName.caseInsensitiveContains(".new.exe", True) Then
                        Dim mainEXEName As String = currentProcessFileName.caseInsensitiveReplace(".new.exe", "")
                        searchForProcessAndKillIt(mainEXEName, False)

                        IO.File.Delete(mainEXEName)
                        IO.File.Copy(currentProcessFileName, mainEXEName)

                        Process.Start(New ProcessStartInfo With {.FileName = mainEXEName})
                        Process.GetCurrentProcess.Kill()
                    Else
                        MsgBox("The environment is not ready for an update. This process will now terminate.", MsgBoxStyle.Critical, "Add Adobe Flash to Microsoft EMET")
                        Process.GetCurrentProcess.Kill()
                    End If
                End If

                e.Cancel = True
                Exit Sub
            End If
        End Sub
    End Class
End Namespace