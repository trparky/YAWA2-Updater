Imports Microsoft.Win32
Imports Microsoft.Win32.TaskScheduler
Imports System.Text

Public Class Form1
    Private strLocationOfCCleaner, remoteINIFileVersion, localINIFileVersion As String
    Private boolWinXP As Boolean = False
    Private Const updateNeeded As String = "Update Needed"
    Private Const updateNotNeeded As String = "Update NOT Needed"

    Sub addTask(taskName As String, taskDescription As String, taskEXEPath As String, taskParameters As String)
        taskName = taskName.Trim
        taskDescription = taskDescription.Trim
        taskEXEPath = taskEXEPath.Trim
        taskParameters = taskParameters.Trim

        If Not IO.File.Exists(taskEXEPath) Then
            MsgBox("Executable path not found.", MsgBoxStyle.Critical, Me.Text)
            Exit Sub
        End If

        Dim taskService As TaskService = New TaskService()
        Dim newTask As TaskDefinition = taskService.NewTask

        newTask.RegistrationInfo.Description = taskDescription
        newTask.Triggers.Add(New TaskScheduler.LogonTrigger)

        Dim exeFileInfo As New IO.FileInfo(taskEXEPath)

        newTask.Actions.Add(New ExecAction(Chr(34) & taskEXEPath & Chr(34), taskParameters, exeFileInfo.DirectoryName))

        newTask.Principal.RunLevel = TaskRunLevel.Highest
        newTask.Settings.Compatibility = TaskCompatibility.V2_1
        newTask.Settings.AllowDemandStart = True
        newTask.Settings.DisallowStartIfOnBatteries = False
        newTask.Settings.RunOnlyIfIdle = False
        newTask.Settings.StopIfGoingOnBatteries = False
        newTask.Settings.AllowHardTerminate = False
        newTask.Settings.UseUnifiedSchedulingEngine = True
        newTask.Settings.ExecutionTimeLimit = Nothing
        newTask.Principal.LogonType = TaskLogonType.InteractiveToken

        taskService.RootFolder.RegisterTaskDefinition(taskName, newTask)

        newTask.Dispose()
        taskService.Dispose()
        newTask = Nothing
        taskService = Nothing
    End Sub

    Function doesTaskExist(nameOfTask As String, ByRef taskObject As Task) As Boolean
        Using taskServiceObject As TaskService = New TaskService
            taskObject = taskServiceObject.GetTask(nameOfTask)

            If taskObject Is Nothing Then
                Return False
            Else
                Return True
            End If
        End Using
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            If programFunctions.areWeAnAdministrator() Then
                lblAdmin.Visible = True
                uacImage.Visible = True
                chkLoadAtUserStartup.Enabled = False
            End If

            newFileDeleter()

            If IO.File.Exists("winapp.ini updater custom entries.txt") Then
                IO.File.Move("winapp.ini updater custom entries.txt", programConstants.customEntriesFile)
            End If

            If Environment.OSVersion.ToString.Contains("5.1") Or Environment.OSVersion.ToString.Contains("5.2") Then
                boolWinXP = True
            End If

            If Not boolWinXP And programFunctions.areWeAnAdministrator() Then
                Dim taskService As New TaskService
                Dim taskObject As Task = Nothing

                chkLoadAtUserStartup.Enabled = True

                If doesTaskExist("WinApp.ini Updater", taskObject) Then
                    addTask("YAWA2 Updater (User " & Environment.UserName & ")", "Updates the WinApp2.ini file for CCleaner at User Logon in Silent Mode", Application.ExecutablePath, "-silent")
                    taskService.RootFolder.DeleteTask("WinApp.ini Updater")
                    chkLoadAtUserStartup.Checked = True
                ElseIf doesTaskExist("YAWA2 Updater (User " & Environment.UserName & ")", taskObject) Then
                    chkLoadAtUserStartup.Checked = True
                End If

                taskService.Dispose()
                taskService = Nothing
            ElseIf Not boolWinXP And Not programFunctions.areWeAnAdministrator() Then
                chkLoadAtUserStartup.Enabled = False
            Else
                chkLoadAtUserStartup.Visible = False
            End If

            Control.CheckForIllegalCrossThreadCalls = False
            chkTrim.Checked = programVariables.boolTrim
            chkNotifyAfterUpdateatLogon.Checked = programVariables.boolNotifyAfterUpdateAtLogon
            chkMobileMode.Checked = programVariables.boolMobileMode

            strLocationOfCCleaner = getLocationOfCCleaner()

            If IO.File.Exists(IO.Path.Combine(strLocationOfCCleaner, "winapp2.ini")) Then
                Dim streamReader As New IO.StreamReader(IO.Path.Combine(strLocationOfCCleaner, "winapp2.ini"))
                localINIFileVersion = programFunctions.getINIVersionFromString(streamReader.ReadLine)
                streamReader.Close()
                streamReader.Dispose()
                streamReader = Nothing
            Else
                localINIFileVersion = "(Not Installed)"
            End If

            lblYourVersion.Text &= " " & localINIFileVersion

            If IO.File.Exists(programConstants.customEntriesFile) Then
                Dim customEntriesFileReader As New IO.StreamReader(programConstants.customEntriesFile)
                txtCustomEntries.Text = customEntriesFileReader.ReadToEnd.Trim
                customEntriesFileReader.Close()
                customEntriesFileReader.Dispose()
                customEntriesFileReader = Nothing
                programFunctions.saveSettingToINIFile(programConstants.configINICustomEntriesKey, programFunctions.convertToBase64(txtCustomEntries.Text))
                IO.File.Delete(programConstants.customEntriesFile)
            Else
                Dim customINIFileEntries As String = Nothing ' This variable is used in a ByRef method later in this routine.

                ' Check if the setting in the INI file exists.
                If programFunctions.loadSettingFromINIFile(programConstants.configINICustomEntriesKey, customINIFileEntries) Then
                    ' Yes, it does; let's work with it.
                    If programFunctions.isBase64(customINIFileEntries) Then ' Checks to see if the data from the INI file is valid Base64 encoded data.
                        customINIFileEntries = programFunctions.convertFromBase64(customINIFileEntries) ' Yes it is... so let's decode the Base64 data back into human readable data.
                        txtCustomEntries.Text = customINIFileEntries ' Put the data into the GUI.
                        customINIFileEntries = Nothing ' Free up memory.
                    Else
                        programFunctions.removeSettingFromINIFile(programConstants.configINICustomEntriesKey) ' Remove the offending data from the INI file.
                        MsgBox("Invalid Base64 encoded data found in custom entries key in config INI file. The invalid data has been removed.", MsgBoxStyle.Information, Me.Text) ' Tell the user that bad data was found and that it has been removed from the INI file.
                    End If
                End If
            End If

            If Environment.OSVersion.ToString.Contains("5.1") Or Environment.OSVersion.ToString.Contains("5.2") Then
                boolWinXP = True
            End If

            If Not boolWinXP Then
                Dim strUseSSL As String = Nothing

                If programFunctions.loadSettingFromINIFile("useSSL", strUseSSL) Then
                    If strUseSSL.Equals("True", StringComparison.OrdinalIgnoreCase) Then
                        boolUseSSL = True
                        chkUseSSL.Checked = True
                    Else
                        boolUseSSL = False
                    End If
                Else
                    programFunctions.saveSettingToINIFile("useSSL", "True")
                    boolUseSSL = True
                End If
            Else
                boolUseSSL = False
                chkUseSSL.Visible = False
            End If

            Threading.ThreadPool.QueueUserWorkItem(AddressOf getINIVersion)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Function getLocationOfCCleaner() As String
        Dim msgBoxResult As MsgBoxResult

        If programVariables.boolMobileMode Then
            'strLocationOfCCleaner = New IO.FileInfo(Application.ExecutablePath).DirectoryName
            Return New IO.FileInfo(Application.ExecutablePath).DirectoryName
        Else
            If Not programFunctions.areWeAnAdministrator() And Short.Parse(Environment.OSVersion.Version.Major) >= 6 Then
                Dim startInfo As New ProcessStartInfo With {
                    .FileName = Process.GetCurrentProcess.MainModule.FileName,
                    .Verb = "runas"
                }
                Process.Start(startInfo)
                Process.GetCurrentProcess.Kill()
                Return Nothing
            End If

            Try
                If Environment.Is64BitOperatingSystem Then
                    If RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey("SOFTWARE\Piriform\CCleaner", False) Is Nothing Then
                        msgBoxResult = MsgBox("CCleaner doesn't appear to be installed on your machine." & vbCrLf & vbCrLf & "Should mobile mode be enabled?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, Me.Text)

                        If msgBoxResult = Microsoft.VisualBasic.MsgBoxResult.Yes Then
                            programVariables.boolMobileMode = True
                            programFunctions.saveSettingToINIFile(programConstants.configINIMobileModeKey, "True")
                            chkMobileMode.Checked = True
                            Return New IO.FileInfo(Application.ExecutablePath).DirectoryName
                        Else
                            Me.Close()
                            Return Nothing
                        End If
                    Else
                        'strLocationOfCCleaner = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey("SOFTWARE\Piriform\CCleaner", False).GetValue(vbNullString, Nothing)
                        Return RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey("SOFTWARE\Piriform\CCleaner", False).GetValue(vbNullString, Nothing)
                    End If
                Else
                    If Registry.LocalMachine.OpenSubKey("SOFTWARE\Piriform\CCleaner", False) Is Nothing Then
                        msgBoxResult = MsgBox("CCleaner doesn't appear to be installed on your machine." & vbCrLf & vbCrLf & "Should mobile mode be enabled?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, Me.Text)

                        If msgBoxResult = Microsoft.VisualBasic.MsgBoxResult.Yes Then
                            programVariables.boolMobileMode = True
                            programFunctions.saveSettingToINIFile(programConstants.configINIMobileModeKey, "True")
                            chkMobileMode.Checked = True
                            Return New IO.FileInfo(Application.ExecutablePath).DirectoryName
                        Else
                            Me.Close()
                            Return Nothing
                        End If
                    Else
                        'strLocationOfCCleaner = Registry.LocalMachine.OpenSubKey("SOFTWARE\Piriform\CCleaner", False).GetValue(vbNullString, Nothing)
                        Return Registry.LocalMachine.OpenSubKey("SOFTWARE\Piriform\CCleaner", False).GetValue(vbNullString, Nothing)
                    End If
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
                Return Nothing
            End Try
        End If
    End Function

    Sub getINIVersion()
        Try
            remoteINIFileVersion = programFunctions.getRemoteINIFileVersion()

            If remoteINIFileVersion = programConstants.errorRetrievingRemoteINIFileVersion Then
                MsgBox("Error Retrieving Remote INI File Version.  Please try again.", MsgBoxStyle.Critical, "WinApp.ini Updater")
                Exit Sub
            End If

            lblWebSiteVersion.Text = "Web Site WinApp2.ini Version: " & remoteINIFileVersion

            If localINIFileVersion = "(Not Installed)" Then
                btnApplyNewINIFile.Enabled = True
                lblUpdateNeededOrNot.Text = updateNeeded
                lblUpdateNeededOrNot.Font = New Font(lblUpdateNeededOrNot.Font.FontFamily, lblUpdateNeededOrNot.Font.SizeInPoints, FontStyle.Bold)
                MsgBox("You don't have a CCleaner WinApp2.ini file installed." & vbCrLf & vbCrLf & "Remote INI File Version: " & remoteINIFileVersion, MsgBoxStyle.Information, Me.Text)
            Else
                If remoteINIFileVersion = localINIFileVersion Then
                    btnApplyNewINIFile.Enabled = False
                    lblUpdateNeededOrNot.Text = updateNotNeeded
                    MsgBox("You already have the latest CCleaner INI file version.", MsgBoxStyle.Information, Me.Text)
                Else
                    btnApplyNewINIFile.Enabled = True
                    lblUpdateNeededOrNot.Text = updateNeeded
                    lblUpdateNeededOrNot.Font = New Font(lblUpdateNeededOrNot.Font.FontFamily, lblUpdateNeededOrNot.Font.SizeInPoints, FontStyle.Bold)
                    MsgBox("There is a new version of the CCleaner WinApp2.ini file." & vbCrLf & vbCrLf & "New Remote INI File Version: " & remoteINIFileVersion, MsgBoxStyle.Information, Me.Text)
                End If
            End If
        Catch ex As Threading.ThreadAbortException
        Catch ex2 As Exception
            MsgBox(ex2.Message)
        End Try
    End Sub

    Private Sub btnSaveCustomEntries_Click(sender As Object, e As EventArgs) Handles btnSaveCustomEntries.Click
        If txtCustomEntries.Text.Trim = Nothing Then
            programFunctions.removeSettingFromINIFile(programConstants.configINICustomEntriesKey)
        Else
            programFunctions.saveSettingToINIFile(programConstants.configINICustomEntriesKey, Convert.ToBase64String(Encoding.UTF8.GetBytes(txtCustomEntries.Text)))
        End If

        MsgBox("Your custom entries have been saved.", MsgBoxStyle.Information, Me.Text)
    End Sub

    Private Sub applyNewINIFileSub()
        Dim remoteINIFileData As String = Nothing

        If internetFunctions.createNewHTTPHelperObject().getWebData(programConstants.WinApp2INIFileURL, remoteINIFileData, False) Then
            Dim streamWriter As New IO.StreamWriter(IO.Path.Combine(strLocationOfCCleaner, "winapp2.ini"))
            streamWriter.Write(remoteINIFileData & vbCrLf & txtCustomEntries.Text & vbCrLf)
            streamWriter.Close()
            streamWriter.Dispose()
            streamWriter = Nothing

            lblYourVersion.Text = "Your WinApp2.ini Version: " & remoteINIFileVersion
            lblUpdateNeededOrNot.Text = updateNotNeeded
            lblUpdateNeededOrNot.Font = New Font(lblUpdateNeededOrNot.Font.FontFamily, lblUpdateNeededOrNot.Font.SizeInPoints, FontStyle.Regular)

            If chkTrim.Checked Then
                MsgBox("New CCleaner INI File Saved.  Trimming of INI file will now commence.", MsgBoxStyle.Information, Me.Text)
                programFunctions.trimINIFile(strLocationOfCCleaner, remoteINIFileVersion, False)
            Else
                If programVariables.boolMobileMode Then
                    MsgBox("New CCleaner INI File Saved.", MsgBoxStyle.Information, Me.Text)
                Else
                    Dim msgBoxResult As MsgBoxResult = MsgBox("New CCleaner INI File Saved." & vbCrLf & vbCrLf & "Do you want to run CCleaner now?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "WinApp.ini Updater")

                    If msgBoxResult = Microsoft.VisualBasic.MsgBoxResult.Yes Then
                        If Environment.Is64BitOperatingSystem Then
                            Process.Start(IO.Path.Combine(strLocationOfCCleaner, "CCleaner64.exe"))
                        Else
                            Process.Start(IO.Path.Combine(strLocationOfCCleaner, "CCleaner.exe"))
                        End If
                    End If
                End If
            End If
        Else
            MsgBox("There was an error while downloading the WinApp2.ini file.", MsgBoxStyle.Information, Me.Text)
            Exit Sub
        End If
    End Sub

    Private Sub reDownloadSub()
        Dim remoteINIFileData As String = Nothing

        If internetFunctions.createNewHTTPHelperObject().getWebData(programConstants.WinApp2INIFileURL, remoteINIFileData, False) Then
            Dim streamWriter As New IO.StreamWriter(IO.Path.Combine(strLocationOfCCleaner, "winapp2.ini"))
            streamWriter.Write(remoteINIFileData & vbCrLf & txtCustomEntries.Text & vbCrLf)
            streamWriter.Close()
            streamWriter.Dispose()
            streamWriter = Nothing

            lblYourVersion.Text = "Your WinApp2.ini Version: " & remoteINIFileVersion

            If chkTrim.Checked Then
                MsgBox("New CCleaner WinApp2.ini File Saved.  Trimming of INI file will now commence.", MsgBoxStyle.Information, Me.Text)
                programFunctions.trimINIFile(strLocationOfCCleaner, remoteINIFileVersion, False)
            Else
                Dim msgBoxResult As MsgBoxResult = MsgBox("New CCleaner WinApp2.ini File Saved." & vbCrLf & vbCrLf & "Do you want to run CCleaner now?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, Me.Text)

                If msgBoxResult = Microsoft.VisualBasic.MsgBoxResult.Yes Then
                    If Environment.Is64BitOperatingSystem Then
                        Process.Start(IO.Path.Combine(strLocationOfCCleaner, "CCleaner64.exe"))
                    Else
                        Process.Start(IO.Path.Combine(strLocationOfCCleaner, "CCleaner.exe"))
                    End If
                End If
            End If
        Else
            MsgBox("There was an error while downloading the WinApp2.ini file.", MsgBoxStyle.Information, Me.Text)
            Exit Sub
        End If
    End Sub

    Private Sub btnTrim_Click(sender As Object, e As EventArgs) Handles btnTrim.Click
        Threading.ThreadPool.QueueUserWorkItem(Sub()
                                                   programFunctions.trimINIFile(strLocationOfCCleaner, remoteINIFileVersion, False)
                                               End Sub)
    End Sub

    Private Sub btnApplyNewINIFile_Click(sender As Object, e As EventArgs) Handles btnApplyNewINIFile.Click
        Threading.ThreadPool.QueueUserWorkItem(AddressOf applyNewINIFileSub)
    End Sub

    Private Sub btnReDownload_Click(sender As Object, e As EventArgs) Handles btnReDownload.Click
        Threading.ThreadPool.QueueUserWorkItem(AddressOf reDownloadSub)
    End Sub

    Private Sub chkLoadAtUserStartup_Click(sender As Object, e As EventArgs) Handles chkLoadAtUserStartup.Click
        If chkLoadAtUserStartup.Checked Then
            addTask("YAWA2 Updater (User " & Environment.UserName & ")", "Updates the WinApp2.ini file for CCleaner at User Logon in Silent Mode", Application.ExecutablePath, "-silent")
        Else
            Dim taskService As New TaskService
            taskService.RootFolder.DeleteTask("YAWA2 Updater")
            taskService.Dispose()
            taskService = Nothing
        End If
    End Sub

    ' ============================
    ' ==== Updater Code Below ====
    ' ============================

    Private Sub btnCheckForUpdates_Click(sender As Object, e As EventArgs) Handles btnCheckForUpdates.Click
        Threading.ThreadPool.QueueUserWorkItem(Sub() checkForUpdates(Me))
        btnCheckForUpdates.Enabled = False
    End Sub

    Public Function checkForInternetConnection() As Boolean
        If My.Computer.Network.IsAvailable Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub btnAbout_Click(sender As Object, e As EventArgs) Handles btnAbout.Click
        Dim version() As String = Application.ProductVersion.Split(".".ToCharArray) ' Gets the program version
        Dim stringBuilder As New Text.StringBuilder

        stringBuilder.AppendLine(Me.Text)
        stringBuilder.AppendLine("Written By Tom Parkison")
        stringBuilder.AppendLine("Copyright Thomas Parkison 2012-2015.")
        stringBuilder.AppendLine()
        stringBuilder.AppendFormat("Version {0}.{1} Build {2}", version(0), version(1), version(2))

        MsgBox(stringBuilder.ToString.Trim, MsgBoxStyle.Information, "About")

        version = Nothing
        stringBuilder = Nothing
    End Sub

    Private Sub chkNotifyAfterUpdateatLogon_Click(sender As Object, e As EventArgs) Handles chkNotifyAfterUpdateatLogon.Click
        programFunctions.saveSettingToINIFile(programConstants.configINInotifyAfterUpdateAtLogonKey, chkNotifyAfterUpdateatLogon.Checked.ToString)
    End Sub

    Private Sub chkMobileMode_Click(sender As Object, e As EventArgs) Handles chkMobileMode.Click
        programFunctions.saveSettingToINIFile(programConstants.configINIMobileModeKey, chkMobileMode.Checked.ToString)
        programVariables.boolMobileMode = chkMobileMode.Checked

        If chkMobileMode.Checked Then
            chkLoadAtUserStartup.Enabled = False
        Else
            chkLoadAtUserStartup.Enabled = True
        End If

        getLocationOfCCleaner()
    End Sub

    Private Sub chkTrim_Click(sender As Object, e As EventArgs) Handles chkTrim.Click
        programFunctions.saveSettingToINIFile(programConstants.configINITrimKey, chkTrim.Checked.ToString)
    End Sub

    Private Sub chkUseSSL_Click(sender As Object, e As EventArgs) Handles chkUseSSL.Click
        If chkUseSSL.Checked Then
            programFunctions.saveSettingToINIFile("useSSL", "True")
            boolUseSSL = True
        Else
            programFunctions.saveSettingToINIFile("useSSL", "False")
            boolUseSSL = False
        End If
    End Sub
End Class