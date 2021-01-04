﻿Imports Microsoft.Win32
Imports Microsoft.Win32.TaskScheduler
Imports System.Text
Imports System.Xml.Serialization

Public Class Form1
    Private strLocationOfCCleaner, remoteINIFileVersion, localINIFileVersion As String
    Private Const updateNeeded As String = "Update Needed"
    Private Const updateNotNeeded As String = "Update NOT Needed"
    Private Const strMessageBoxTitle As String = "WinApp.ini Updater"
    Private boolDoneLoading As Boolean = False

    Sub AddTask(taskName As String, taskDescription As String, taskEXEPath As String, taskParameters As String)
        taskName = taskName.Trim
        taskDescription = taskDescription.Trim
        taskEXEPath = taskEXEPath.Trim
        taskParameters = taskParameters.Trim

        If Not IO.File.Exists(taskEXEPath) Then
            WPFCustomMessageBox.CustomMessageBox.ShowOK("Executable path not found.", strMessageBoxTitle, programConstants.strOK, Windows.MessageBoxImage.Error)
            Exit Sub
        End If

        Dim taskService As TaskService = New TaskService()
        Dim newTask As TaskDefinition = taskService.NewTask
        Dim exeFileInfo As New IO.FileInfo(taskEXEPath)

        With newTask
            .RegistrationInfo.Description = taskDescription
            .Triggers.Add(New LogonTrigger)
            .Actions.Add(New ExecAction(Chr(34) & taskEXEPath & Chr(34), taskParameters, exeFileInfo.DirectoryName))

            .Principal.RunLevel = TaskRunLevel.Highest
            .Settings.Compatibility = TaskCompatibility.V2_1
            .Settings.AllowDemandStart = True
            .Settings.DisallowStartIfOnBatteries = False
            .Settings.RunOnlyIfIdle = False
            .Settings.StopIfGoingOnBatteries = False
            .Settings.AllowHardTerminate = False
            .Settings.UseUnifiedSchedulingEngine = True
            .Settings.ExecutionTimeLimit = Nothing
            .Principal.LogonType = TaskLogonType.InteractiveToken
        End With

        taskService.RootFolder.RegisterTaskDefinition(taskName, newTask)

        newTask.Dispose()
        taskService.Dispose()
    End Sub

    Function DoesTaskExist(nameOfTask As String, ByRef taskObject As Task) As Boolean
        Using taskServiceObject As TaskService = New TaskService
            taskObject = taskServiceObject.GetTask(nameOfTask)
            Return taskObject IsNot Nothing
        End Using
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Dim boolAreWeAnAdmin As Boolean = programFunctions.AreWeAnAdministrator()
            lblAdmin.Visible = boolAreWeAnAdmin
            chkLoadAtUserStartup.Enabled = Not boolAreWeAnAdmin

            NewFileDeleter()

            If programFunctions.AreWeAnAdministrator() Then
                Dim taskService As New TaskService
                Dim taskObject As Task = Nothing

                chkLoadAtUserStartup.Enabled = True

                If DoesTaskExist("WinApp.ini Updater", taskObject) Then
                    AddTask("YAWA2 Updater (User " & Environment.UserName & ")", "Updates the WinApp2.ini file for CCleaner at User Logon in Silent Mode", Application.ExecutablePath, "-silent")
                    taskService.RootFolder.DeleteTask("WinApp.ini Updater")
                    chkLoadAtUserStartup.Checked = True
                ElseIf DoesTaskExist("YAWA2 Updater (User " & Environment.UserName & ")", taskObject) Then
                    chkLoadAtUserStartup.Checked = True
                End If

                taskService.Dispose()
                taskService = Nothing
            ElseIf Not programFunctions.AreWeAnAdministrator() Then : chkLoadAtUserStartup.Enabled = False
            Else : chkLoadAtUserStartup.Visible = False
            End If

            Dim AppSettings As New AppSettings
            SyncLock programFunctions.LockObject
                Using streamReader As New IO.StreamReader(programConstants.configXMLFile)
                    Dim xmlSerializerObject As New XmlSerializer(AppSettings.GetType)
                    AppSettings = xmlSerializerObject.Deserialize(streamReader)
                End Using
            End SyncLock

            If Not String.IsNullOrEmpty(AppSettings.strCustomEntries) Then TxtCustomEntries.Text = AppSettings.strCustomEntries.Replace(vbLf, vbCrLf)
            chkTrim.Checked = AppSettings.boolTrim
            chkNotifyAfterUpdateatLogon.Checked = AppSettings.boolNotifyAfterUpdateAtLogon
            chkMobileMode.Checked = AppSettings.boolMobileMode
            chkUseSSL.Checked = AppSettings.boolUseSSL
            ChkSleepOnSilentStartup.Checked = AppSettings.boolSleepOnSilentStartup

            strLocationOfCCleaner = GetLocationOfCCleaner()

            If IO.File.Exists(IO.Path.Combine(strLocationOfCCleaner, "winapp2.ini")) Then
                Using streamReader As New IO.StreamReader(IO.Path.Combine(strLocationOfCCleaner, "winapp2.ini"))
                    localINIFileVersion = programFunctions.GetINIVersionFromString(streamReader.ReadLine)
                End Using
            Else : localINIFileVersion = "(Not Installed)"
            End If

            lblYourVersion.Text &= " " & localINIFileVersion

            Threading.ThreadPool.QueueUserWorkItem(AddressOf GetINIVersion)
            boolDoneLoading = True
        Catch ex As Exception
            WPFCustomMessageBox.CustomMessageBox.ShowOK(ex.Message, strMessageBoxTitle, programConstants.strOK)
        End Try
    End Sub

    Function GetLocationOfCCleaner() As String
        Dim msgBoxResult As Windows.MessageBoxResult

        If chkMobileMode.Checked Then
            'strLocationOfCCleaner = New IO.FileInfo(Application.ExecutablePath).DirectoryName
            Return New IO.FileInfo(Application.ExecutablePath).DirectoryName
        Else
            If Not programFunctions.AreWeAnAdministrator() And Short.Parse(Environment.OSVersion.Version.Major) >= 6 Then
                Dim startInfo As New ProcessStartInfo With {.FileName = Process.GetCurrentProcess.MainModule.FileName, .Verb = "runas"}
                Process.Start(startInfo)
                Process.GetCurrentProcess.Kill()
                Return Nothing
            End If

            Try
                If Environment.Is64BitOperatingSystem Then
                    If RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey("SOFTWARE\Piriform\CCleaner", False) Is Nothing Then
                        msgBoxResult = WPFCustomMessageBox.CustomMessageBox.ShowYesNo("CCleaner doesn't appear to be installed on your machine." & vbCrLf & vbCrLf & "Should mobile mode be enabled?", strMessageBoxTitle, programConstants.strYes, programConstants.strNo, Windows.MessageBoxImage.Question)

                        If msgBoxResult = Windows.MessageBoxResult.Yes Then
                            programFunctions.SaveSettingToAppSettingsXMLFile(programFunctions.AppSettingType.boolMobileMode, True)
                            chkMobileMode.Checked = True
                            Return New IO.FileInfo(Application.ExecutablePath).DirectoryName
                        Else
                            Me.Close()
                            Return Nothing
                        End If
                    Else
                        Return RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey("SOFTWARE\Piriform\CCleaner", False).GetValue(vbNullString, Nothing)
                    End If
                Else
                    If Registry.LocalMachine.OpenSubKey("SOFTWARE\Piriform\CCleaner", False) Is Nothing Then
                        msgBoxResult = WPFCustomMessageBox.CustomMessageBox.ShowYesNo("CCleaner doesn't appear to be installed on your machine." & vbCrLf & vbCrLf & "Should mobile mode be enabled?", strMessageBoxTitle, programConstants.strYes, programConstants.strNo, Windows.MessageBoxImage.Question)

                        If msgBoxResult = Windows.MessageBoxResult.Yes Then
                            programFunctions.SaveSettingToAppSettingsXMLFile(programFunctions.AppSettingType.boolMobileMode, True)
                            chkMobileMode.Checked = True
                            Return New IO.FileInfo(Application.ExecutablePath).DirectoryName
                        Else
                            Me.Close()
                            Return Nothing
                        End If
                    Else
                        Return Registry.LocalMachine.OpenSubKey("SOFTWARE\Piriform\CCleaner", False).GetValue(vbNullString, Nothing)
                    End If
                End If
            Catch ex As Exception
                WPFCustomMessageBox.CustomMessageBox.ShowOK(ex.Message, strMessageBoxTitle, programConstants.strOK)
                Return Nothing
            End Try
        End If
    End Function

    Sub GetINIVersion()
        Try
            remoteINIFileVersion = programFunctions.GetRemoteINIFileVersion()

            If remoteINIFileVersion = programConstants.errorRetrievingRemoteINIFileVersion Then
                WPFCustomMessageBox.CustomMessageBox.ShowOK("Error Retrieving Remote INI File Version.  Please try again.", strMessageBoxTitle, programConstants.strOK, Windows.MessageBoxImage.Error)
                Exit Sub
            End If

            Me.Invoke(Sub()
                          lblWebSiteVersion.Text = "Web Site WinApp2.ini Version: " & remoteINIFileVersion

                          If localINIFileVersion = "(Not Installed)" Then
                              btnApplyNewINIFile.Enabled = True
                              lblUpdateNeededOrNot.Text = updateNeeded
                              lblUpdateNeededOrNot.Font = New Font(lblUpdateNeededOrNot.Font.FontFamily, lblUpdateNeededOrNot.Font.SizeInPoints, FontStyle.Bold)
                              WPFCustomMessageBox.CustomMessageBox.ShowOK("You don't have a CCleaner WinApp2.ini file installed." & vbCrLf & vbCrLf & "Remote INI File Version: " & remoteINIFileVersion, strMessageBoxTitle, programConstants.strOK, Windows.MessageBoxImage.Information)
                          Else
                              If remoteINIFileVersion = localINIFileVersion Then
                                  btnApplyNewINIFile.Enabled = False
                                  lblUpdateNeededOrNot.Text = updateNotNeeded
                                  WPFCustomMessageBox.CustomMessageBox.ShowOK("You already have the latest CCleaner INI file version.", strMessageBoxTitle, programConstants.strOK, Windows.MessageBoxImage.Information)
                              Else
                                  btnApplyNewINIFile.Enabled = True
                                  lblUpdateNeededOrNot.Text = updateNeeded
                                  lblUpdateNeededOrNot.Font = New Font(lblUpdateNeededOrNot.Font.FontFamily, lblUpdateNeededOrNot.Font.SizeInPoints, FontStyle.Bold)

                                  Dim stringBuilder As New StringBuilder()
                                  stringBuilder.AppendLine("There is a new version of the CCleaner WinApp2.ini file.")
                                  stringBuilder.AppendLine()
                                  stringBuilder.AppendLine("Currently Installed INI File Version: " & localINIFileVersion)
                                  stringBuilder.AppendLine("New Remote INI File Version: " & remoteINIFileVersion)

                                  WPFCustomMessageBox.CustomMessageBox.ShowOK(stringBuilder.ToString.Trim, strMessageBoxTitle, programConstants.strOK, Windows.MessageBoxImage.Information)
                              End If
                          End If
                      End Sub)
        Catch ex As Threading.ThreadAbortException
        Catch ex2 As Exception
            WPFCustomMessageBox.CustomMessageBox.ShowOK(ex2.Message, strMessageBoxTitle, programConstants.strOK)
        End Try
    End Sub

    Private Sub DownloadINIFileAndSaveIt(Optional boolUpdateLabelOnGUI As Boolean = False)
        Dim remoteINIFileData As String = Nothing

        If internetFunctions.CreateNewHTTPHelperObject().GetWebData(programConstants.WinApp2INIFileURL, remoteINIFileData, False) Then
            Using streamWriter As New IO.StreamWriter(IO.Path.Combine(strLocationOfCCleaner, "winapp2.ini"))
                streamWriter.Write(If(String.IsNullOrEmpty(TxtCustomEntries.Text), remoteINIFileData & vbCrLf, remoteINIFileData & vbCrLf & TxtCustomEntries.Text & vbCrLf))
            End Using

            Me.Invoke(Sub()
                          lblYourVersion.Text = "Your WinApp2.ini Version: " & remoteINIFileVersion

                          If boolUpdateLabelOnGUI Then
                              lblUpdateNeededOrNot.Text = updateNotNeeded
                              lblUpdateNeededOrNot.Font = New Font(lblUpdateNeededOrNot.Font.FontFamily, lblUpdateNeededOrNot.Font.SizeInPoints, FontStyle.Regular)
                          End If

                          If chkTrim.Checked Then
                              WPFCustomMessageBox.CustomMessageBox.ShowOK("New CCleaner WinApp2.ini File Saved. Trimming of INI file will now commence.", strMessageBoxTitle, programConstants.strOK, Windows.MessageBoxImage.Information)
                              programFunctions.TrimINIFile(strLocationOfCCleaner, remoteINIFileVersion, False)
                          Else
                              If chkMobileMode.Checked Then
                                  WPFCustomMessageBox.CustomMessageBox.ShowOK("New CCleaner WinApp2.ini File Saved.", strMessageBoxTitle, programConstants.strOK, Windows.MessageBoxImage.Information)
                              Else
                                  If WPFCustomMessageBox.CustomMessageBox.ShowYesNo("New CCleaner WinApp2.ini File Saved." & vbCrLf & vbCrLf & "Do you want to run CCleaner now?", strMessageBoxTitle, programConstants.strYes, programConstants.strNo, Windows.MessageBoxImage.Question) = Windows.MessageBoxResult.Yes Then programFunctions.RunCCleaner(strLocationOfCCleaner)
                              End If
                          End If
                      End Sub)
        Else
            WPFCustomMessageBox.CustomMessageBox.ShowOK("There was an error while downloading the WinApp2.ini file.", strMessageBoxTitle, programConstants.strOK, Windows.MessageBoxImage.Information)
        End If
    End Sub

    Private Sub BtnTrim_Click(sender As Object, e As EventArgs) Handles btnTrim.Click
        Threading.ThreadPool.QueueUserWorkItem(Sub()
                                                   programFunctions.TrimINIFile(strLocationOfCCleaner, remoteINIFileVersion, False)
                                               End Sub)
    End Sub

    Private Sub BtnApplyNewINIFile_Click(sender As Object, e As EventArgs) Handles btnApplyNewINIFile.Click
        Threading.ThreadPool.QueueUserWorkItem(Sub() DownloadINIFileAndSaveIt(True))
    End Sub

    Private Sub BtnReDownload_Click(sender As Object, e As EventArgs) Handles btnReDownload.Click
        Threading.ThreadPool.QueueUserWorkItem(Sub() DownloadINIFileAndSaveIt())
    End Sub

    Private Sub ChkLoadAtUserStartup_Click(sender As Object, e As EventArgs) Handles chkLoadAtUserStartup.Click
        If chkLoadAtUserStartup.Checked Then
            AddTask("YAWA2 Updater (User " & Environment.UserName & ")", "Updates the WinApp2.ini file for CCleaner at User Logon in Silent Mode", Application.ExecutablePath, "-silent")
        Else
            Using taskService As New TaskService
                taskService.RootFolder.DeleteTask("YAWA2 Updater")
            End Using
        End If
    End Sub

    ' ============================
    ' ==== Updater Code Below ====
    ' ============================

    Private Sub BtnCheckForUpdates_Click(sender As Object, e As EventArgs) Handles btnCheckForUpdates.Click
        Threading.ThreadPool.QueueUserWorkItem(Sub()
                                                   Dim checkForUpdatesClassObject As New Check_for_Update_Stuff(Me)
                                                   checkForUpdatesClassObject.CheckForUpdates()
                                               End Sub)
        btnCheckForUpdates.Enabled = False
    End Sub

    Public Function CheckForInternetConnection() As Boolean
        Return My.Computer.Network.IsAvailable
    End Function

    Private Sub BtnAbout_Click(sender As Object, e As EventArgs) Handles btnAbout.Click
        Dim version() As String = Application.ProductVersion.Split(".".ToCharArray) ' Gets the program version
        Dim stringBuilder As New Text.StringBuilder

        stringBuilder.AppendLine("WinApp.ini Updater")
        stringBuilder.AppendLine("Written By Tom Parkison")
        stringBuilder.AppendLine("Copyright Thomas Parkison 2012-2020.")
        stringBuilder.AppendLine()
        stringBuilder.AppendFormat("Version {0}.{1} Build {2}", version(0), version(1), version(2))

        WPFCustomMessageBox.CustomMessageBox.ShowOK(stringBuilder.ToString.Trim, "About", programConstants.strOK, Windows.MessageBoxImage.Information)
    End Sub

    Private Sub ChkNotifyAfterUpdateatLogon_Click(sender As Object, e As EventArgs) Handles chkNotifyAfterUpdateatLogon.Click
        programFunctions.SaveSettingToAppSettingsXMLFile(programFunctions.AppSettingType.boolNotifyAfterUpdateAtLogon, chkNotifyAfterUpdateatLogon.Checked)
    End Sub

    Private Sub ChkMobileMode_Click(sender As Object, e As EventArgs) Handles chkMobileMode.Click
        programFunctions.SaveSettingToAppSettingsXMLFile(programFunctions.AppSettingType.boolMobileMode, chkMobileMode.Checked)
        chkLoadAtUserStartup.Enabled = Not chkMobileMode.Checked
        GetLocationOfCCleaner()
    End Sub

    Private Sub ChkTrim_Click(sender As Object, e As EventArgs) Handles chkTrim.Click
        programFunctions.SaveSettingToAppSettingsXMLFile(programFunctions.AppSettingType.boolTrim, chkTrim.Checked)
    End Sub

    Private Sub ChkUseSSL_Click(sender As Object, e As EventArgs) Handles chkUseSSL.Click
        programFunctions.SaveSettingToAppSettingsXMLFile(programFunctions.AppSettingType.boolUseSSL, chkUseSSL.Checked)
    End Sub

    Private Sub TxtCustomEntries_TextChanged(sender As Object, e As EventArgs) Handles TxtCustomEntries.TextChanged
        If boolDoneLoading Then
            TxtCustomEntries.Text = TxtCustomEntries.Text
            programFunctions.SaveCustomEntriesToAppSettingsXMLFile(TxtCustomEntries.Text)
        End If
    End Sub

    Private Sub ChkSleepOnSilentStartup_Click(sender As Object, e As EventArgs) Handles ChkSleepOnSilentStartup.Click
        programFunctions.SaveSettingToAppSettingsXMLFile(programFunctions.AppSettingType.boolSleepOnSilentStartup, ChkSleepOnSilentStartup.Checked)
    End Sub
End Class