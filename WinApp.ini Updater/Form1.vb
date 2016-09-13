Imports Microsoft.Win32
Imports System.Text.RegularExpressions
Imports Microsoft.Win32.TaskScheduler
Imports System.Management

Public Class Form1
    Private strLocationOfCCleaner, remoteINIFileVersion, localINIFileVersion As String
    Private boolWinXP As Boolean = False
    Private Const updateNeeded As String = "Update Needed"
    Private Const updateNotNeeded As String = "Update NOT Needed"

    'Sub addGenericAdminTask(taskName As String, taskDescription As String, taskEXEPath As String, taskParameters As String)
    '    taskName = taskName.Trim
    '    taskDescription = taskDescription.Trim
    '    taskEXEPath = taskEXEPath.Trim
    '    taskParameters = taskParameters.Trim

    '    If System.IO.File.Exists(taskEXEPath) = False Then
    '        MsgBox("Executable path not found.", MsgBoxStyle.Critical, Me.Text)
    '        Exit Sub
    '    End If

    '    Dim taskService As TaskService = New TaskService()
    '    Dim newTask As TaskDefinition = taskService.NewTask

    '    newTask.RegistrationInfo.Description = taskDescription
    '    newTask.Triggers.Add(New Microsoft.Win32.TaskScheduler.LogonTrigger)

    '    Dim exeFileInfo As New System.IO.FileInfo(taskEXEPath)

    '    newTask.Actions.Add(New ExecAction(Chr(34) & taskEXEPath & Chr(34), taskParameters, exeFileInfo.DirectoryName))

    '    'If parameters = Nothing Then
    '    '    newTask.Actions.Add(New ExecAction(Chr(34) & txtEXEPath.Text & Chr(34), Nothing, exeFileInfo.DirectoryName))
    '    'Else
    '    '    newTask.Actions.Add(New ExecAction(Chr(34) & txtEXEPath.Text & Chr(34), parameters, exeFileInfo.DirectoryName))
    '    'End If

    '    newTask.Principal.RunLevel = TaskRunLevel.Highest
    '    newTask.Principal.GroupId = "BUILTIN\Administrators"
    '    newTask.Settings.Compatibility = TaskCompatibility.V2_1
    '    newTask.Settings.AllowDemandStart = True
    '    newTask.Settings.DisallowStartIfOnBatteries = False
    '    newTask.Settings.RunOnlyIfIdle = False
    '    newTask.Settings.StopIfGoingOnBatteries = False
    '    newTask.Settings.AllowHardTerminate = False
    '    newTask.Settings.UseUnifiedSchedulingEngine = True
    '    newTask.Settings.ExecutionTimeLimit = Nothing
    '    newTask.Principal.LogonType = TaskLogonType.InteractiveToken

    '    taskService.RootFolder.RegisterTaskDefinition(taskName, newTask)

    '    newTask.Dispose()
    '    taskService.Dispose()
    '    newTask = Nothing
    '    taskService = Nothing
    'End Sub

    Sub addTask(taskName As String, taskDescription As String, taskEXEPath As String, taskParameters As String)
        taskName = taskName.Trim
        taskDescription = taskDescription.Trim
        taskEXEPath = taskEXEPath.Trim
        taskParameters = taskParameters.Trim

        If System.IO.File.Exists(taskEXEPath) = False Then
            MsgBox("Executable path not found.", MsgBoxStyle.Critical, Me.Text)
            Exit Sub
        End If

        Dim taskService As TaskService = New TaskService()
        Dim newTask As TaskDefinition = taskService.NewTask

        newTask.RegistrationInfo.Description = taskDescription
        newTask.Triggers.Add(New Microsoft.Win32.TaskScheduler.LogonTrigger)

        Dim exeFileInfo As New System.IO.FileInfo(taskEXEPath)

        newTask.Actions.Add(New ExecAction(Chr(34) & taskEXEPath & Chr(34), taskParameters, exeFileInfo.DirectoryName))

        'If parameters = Nothing Then
        '    newTask.Actions.Add(New ExecAction(Chr(34) & txtEXEPath.Text & Chr(34), Nothing, exeFileInfo.DirectoryName))
        'Else
        '    newTask.Actions.Add(New ExecAction(Chr(34) & txtEXEPath.Text & Chr(34), parameters, exeFileInfo.DirectoryName))
        'End If

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

    Private Function doesPIDExist(PID As Integer) As Boolean
        Dim searcher As New ManagementObjectSearcher("root\CIMV2", String.Format("SELECT * FROM Win32_Process WHERE ProcessId={0}", PID))

        If searcher.Get.Count = 0 Then
            searcher.Dispose()
            Return False
        Else
            searcher.Dispose()
            Return True
        End If
    End Function

    Private Sub killProcess(PID As Integer)
        Dim processDetail As Process

        Debug.Write(String.Format("Killing PID {0}...", PID))

        processDetail = Process.GetProcessById(PID)
        processDetail.Kill()

        Threading.Thread.Sleep(100)

        If doesPIDExist(PID) Then
            Debug.WriteLine(" Process still running.  Attempting to kill process again.")
            killProcess(PID)
        Else
            Debug.WriteLine(" Process Killed.")
        End If
    End Sub

    Public Sub searchForProcessAndKillIt(fileName As String)
        Dim fullFileName As String = New IO.FileInfo(fileName).FullName
        'Dim PID As Integer

        Debug.WriteLine("Killing all processes that belong to parent executable file.  Please Wait.")
        'Console.WriteLine(String.Format("SELECT * FROM Win32_Process WHERE ExecutablePath = '{0}'", fullFileName.Replace("\", "\\")))

        Dim searcher As New ManagementObjectSearcher("root\CIMV2", "SELECT * FROM Win32_Process")

        Try
            For Each queryObj As ManagementObject In searcher.Get()
                If queryObj("ExecutablePath") IsNot Nothing Then
                    If queryObj("ExecutablePath") = fullFileName Then
                        killProcess(Integer.Parse(queryObj("ProcessId").ToString))
                    End If
                End If
            Next

            Debug.WriteLine("All processes killed... Update process can continue.")
        Catch err As ManagementException
            ' Does nothing
        End Try
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

            If IO.File.Exists(Application.ExecutablePath & ".new.exe") = True Then
                Dim newFileDeleterThread As New Threading.Thread(Sub()
                                                                     searchForProcessAndKillIt(Application.ExecutablePath & ".new.exe")
                                                                     IO.File.Delete(Application.ExecutablePath & ".new.exe")
                                                                 End Sub)
                newFileDeleterThread.Start()

                'Dim updaterExecutableRemoverThread As New Threading.Thread(AddressOf updaterExecutableRemover)
                'updaterExecutableRemoverThread.Name = "updater.exe Remover Thread"
                'updaterExecutableRemoverThread.Priority = Threading.ThreadPriority.Lowest
                'updaterExecutableRemoverThread.Start()
            End If

            If IO.File.Exists("winapp.ini updater custom entries.txt") Then
                My.Computer.FileSystem.RenameFile("winapp.ini updater custom entries.txt", programConstants.customEntriesFile)
            End If

            If Environment.OSVersion.ToString.Contains("5.1") Or Environment.OSVersion.ToString.Contains("5.2") Then
                boolWinXP = True
            End If

            If boolWinXP = False And programFunctions.areWeAnAdministrator() = True Then
                Dim taskService As New TaskService
                Dim taskObject As Task = Nothing

                chkLoadAtUserStartup.Enabled = True

                If doesTaskExist("WinApp.ini Updater", taskObject) = True Then
                    addTask("YAWA2 Updater (User " & Environment.UserName & ")", "Updates the WinApp2.ini file for CCleaner at User Logon in Silent Mode", Application.ExecutablePath, "-silent")
                    taskService.RootFolder.DeleteTask("WinApp.ini Updater")
                    chkLoadAtUserStartup.Checked = True
                ElseIf doesTaskExist("YAWA2 Updater (User " & Environment.UserName & ")", taskObject) = True Then
                    chkLoadAtUserStartup.Checked = True
                End If

                'For Each task As Task In taskService.RootFolder.Tasks
                '    If task.Name = "WinApp.ini Updater" Then
                '        addTask("YAWA2 Updater", "Updates the WinApp2.ini file for CCleaner at User Logon in Silent Mode", Application.ExecutablePath, "-silent")
                '        taskService.RootFolder.DeleteTask("WinApp.ini Updater")
                '        chkLoadAtUserStartup.Checked = True
                '    ElseIf task.Name = "YAWA2 Updater" Then
                '        chkLoadAtUserStartup.Checked = True
                '    End If
                'Next
                taskService.Dispose()
                taskService = Nothing
            ElseIf boolWinXP = False And programFunctions.areWeAnAdministrator() = False Then
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

            If IO.File.Exists(programConstants.customEntriesFile) = True Then
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
                If programFunctions.loadSettingFromINIFile(programConstants.configINICustomEntriesKey, customINIFileEntries) = True Then
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

            If Environment.OSVersion.ToString.Contains("5.1") = True Or Environment.OSVersion.ToString.Contains("5.2") = True Then
                boolWinXP = True
            End If

            If boolWinXP = False Then
                Dim strUseSSL As String = Nothing

                If programFunctions.loadSettingFromINIFile("useSSL", strUseSSL) = False Then
                    programFunctions.saveSettingToINIFile("useSSL", "True")
                    boolUseSSL = True
                Else
                    If strUseSSL = "True" Then
                        boolUseSSL = True
                        chkUseSSL.Checked = True
                    Else
                        boolUseSSL = False
                    End If
                End If
            Else
                boolUseSSL = False
                chkUseSSL.Visible = False
            End If

            Dim downloadThread As New Threading.Thread(AddressOf getVersionFromWebSite)
            downloadThread.Name = "INI File Downloading Thread"
            downloadThread.Priority = Threading.ThreadPriority.Lowest
            downloadThread.Start()
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
            If programFunctions.areWeAnAdministrator() = False And Short.Parse(Environment.OSVersion.Version.Major) >= 6 Then
                Dim startInfo As New ProcessStartInfo
                startInfo.FileName = Process.GetCurrentProcess.MainModule.FileName
                startInfo.Verb = "runas"
                Process.Start(startInfo)
                Process.GetCurrentProcess.Kill()
                Return Nothing
            End If

            Try
                If Environment.Is64BitOperatingSystem = True Then
                    If RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey("SOFTWARE\Piriform\CCleaner", False) Is Nothing Then
                        msgBoxResult = MsgBox("CCleaner doesn't appear to be installed on your machine." & vbCrLf & vbCrLf & "Should mobile mode be enabled?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, Me.Text)

                        If msgBoxResult = Microsoft.VisualBasic.MsgBoxResult.Yes Then
                            programVariables.boolMobileMode = True
                            programFunctions.saveSettingToINIFile(programConstants.configINIMobileModeKey, "True")
                            'strLocationOfCCleaner = New IO.FileInfo(Application.ExecutablePath).DirectoryName
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
                    If Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\Piriform\CCleaner", False) Is Nothing Then
                        msgBoxResult = MsgBox("CCleaner doesn't appear to be installed on your machine." & vbCrLf & vbCrLf & "Should mobile mode be enabled?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, Me.Text)

                        If msgBoxResult = Microsoft.VisualBasic.MsgBoxResult.Yes Then
                            programVariables.boolMobileMode = True
                            programFunctions.saveSettingToINIFile(programConstants.configINIMobileModeKey, "True")
                            'strLocationOfCCleaner = New IO.FileInfo(Application.ExecutablePath).DirectoryName
                            chkMobileMode.Checked = True
                            Return New IO.FileInfo(Application.ExecutablePath).DirectoryName
                        Else
                            Me.Close()
                            Return Nothing
                        End If
                    Else
                        'strLocationOfCCleaner = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\Piriform\CCleaner", False).GetValue(vbNullString, Nothing)
                        Return Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\Piriform\CCleaner", False).GetValue(vbNullString, Nothing)
                    End If
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
                Return Nothing
            End Try
        End If
    End Function

    Sub getVersionFromWebSite()
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
            programFunctions.saveSettingToINIFile(programConstants.configINICustomEntriesKey, System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(txtCustomEntries.Text)))
        End If

        MsgBox("Your custom entries have been saved.", MsgBoxStyle.Information, Me.Text)
    End Sub

    Private Sub btnApplyNewINIFile_Click(sender As Object, e As EventArgs) Handles btnApplyNewINIFile.Click
        Dim workingThread As New Threading.Thread(Sub()
                                                      Dim httpHelper As httpHelper = internetFunctions.createNewHTTPHelperObject()

                                                      Dim remoteINIFileData As String = Nothing
                                                      If httpHelper.getWebData(programConstants.WinApp2INIFileURL, remoteINIFileData, False) = True Then
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
                                                                      If Environment.Is64BitOperatingSystem = True Then
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
                                                  End Sub)

        workingThread.Start()
    End Sub

    Private Sub btnTrim_Click(sender As Object, e As EventArgs) Handles btnTrim.Click
        Dim trimThread As New Threading.Thread(Sub()
                                                   programFunctions.trimINIFile(strLocationOfCCleaner, remoteINIFileVersion, False)
                                               End Sub)
        trimThread.Name = "INI File Trimming Operation"
        trimThread.Start()
    End Sub

    Private Sub btnReDownload_Click(sender As Object, e As EventArgs) Handles btnReDownload.Click
        Dim workingThread As New Threading.Thread(Sub()
                                                      Dim remoteINIFileData As String = Nothing
                                                      Dim httpHelper As httpHelper = internetFunctions.createNewHTTPHelperObject()
                                                      If httpHelper.getWebData(programConstants.WinApp2INIFileURL, remoteINIFileData, False) = True Then
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
                                                                  If Environment.Is64BitOperatingSystem = True Then
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
                                                  End Sub)

        workingThread.Start()
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
        Dim userInitiatedCheckForUpdatesThread As New Threading.Thread(AddressOf userInitiatedCheckForUpdates)
        userInitiatedCheckForUpdatesThread.Name = "User Initiated Check For Updates Thread"
        userInitiatedCheckForUpdatesThread.Priority = Threading.ThreadPriority.Lowest
        userInitiatedCheckForUpdatesThread.Start()
        btnCheckForUpdates.Enabled = False
    End Sub

    Private Const updaterFileName As String = "updater.exe"

    Private Const programZipFileURL = "www.toms-world.org/download/YAWA2 Updater.zip"
    Private Const programZipFileSHA1URL = "www.toms-world.org/download/YAWA2 Updater.zip.sha1"

    Private Const zipFileName As String = "YAWA2 Updater.zip"
    Private Const programFileNameInZIP As String = "YAWA2 Updater.exe"

    Private Const webSiteURL As String = "www.toms-world.org/blog/yawa2updater"

    Private Const programCodeName As String = "yawa2updater"
    Private Const programUpdateCheckerURLPOST As String = "www.toms-world.org/programupdatechecker"
    'Private Const programUpdateCheckerURL As String = "www.toms-world.org/programupdatechecker/yawa2updater/version/"
    Private Const programName As String = "YAWA2 Updater"

    Sub userInitiatedCheckForUpdates()
        Dim internetConnectionCheckResult As Boolean = checkForInternetConnection()

        If internetConnectionCheckResult = False Then
            btnCheckForUpdates.Enabled = True
            MsgBox("No Internet connection detected.", MsgBoxStyle.Information, Me.Text)
        ElseIf internetConnectionCheckResult = True Then
            Try
                Dim version() As String = Application.ProductVersion.Split(".".ToCharArray) ' Gets the program version

                Dim majorVersion As Short = Short.Parse(version(0))
                Dim minorVersion As Short = Short.Parse(version(1))
                Dim buildVersion As Short = Short.Parse(version(2))

                Dim httpHelper As httpHelper = internetFunctions.createNewHTTPHelperObject()

                httpHelper.addPOSTData("program", "yawa2updater")
                httpHelper.addPOSTData("version", majorVersion & "." & minorVersion)

                Dim strRemoteBuild As String = Nothing, shortRemoteBuild As Short

                If httpHelper.getWebData(programUpdateCheckerURLPOST, strRemoteBuild, False) = True Then
                    Debug.WriteLine("strRemoteBuild = " & strRemoteBuild)

                    ' This handles entirely new versions, not just new builds.
                    If strRemoteBuild.Contains("newversion") = True Then
                        ' Example: newversion-1.2
                        Dim strRemoteBuildParts As String() = strRemoteBuild.Split("-")
                        MsgBox(String.Format("{3} version {0}.{1} is no longer supported and has been replaced by version {2}.", majorVersion, minorVersion, strRemoteBuildParts(1), programName), MsgBoxStyle.Information, Me.Text)
                        downloadAndDoUpdate()
                        Exit Sub
                    ElseIf strRemoteBuild.Contains("beta") = True Then
                        Dim strRemoteBuildParts As String() = strRemoteBuild.Split("-")
                        If strRemoteBuildParts(1) > buildVersion Then
                            Dim updateQuestion As MsgBoxResult = MsgBox("There is an update available but it's classified as a beta/test version." & vbCrLf & vbCrLf & "Do you want to download the update?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, Me.Text)
                            If updateQuestion = MsgBoxResult.Yes Then
                                downloadAndDoUpdate()
                                Exit Sub
                            End If
                        ElseIf Short.Parse(strRemoteBuildParts(1)) = buildVersion Then
                            MsgBox("You already have the latest version.", MsgBoxStyle.Information, Me.Text)
                        End If
                    ElseIf strRemoteBuild.Contains("minor") = True Then
                        Dim strRemoteBuildParts As String() = strRemoteBuild.Split("-")
                        Debug.WriteLine(strRemoteBuildParts.ToString)

                        Dim strRemoteBuildParts2Split As New Specialized.StringCollection
                        If strRemoteBuildParts(2).Contains(",") Then
                            strRemoteBuildParts2Split.AddRange(strRemoteBuildParts(2).ToString.Trim.Split(","))
                            'For Each item As String In strRemoteBuildParts(2).ToString.Trim.Split(",")
                            '    strRemoteBuildParts2Split.Add(item.Trim)
                            'Next
                        Else
                            strRemoteBuildParts2Split.Add(strRemoteBuildParts(2).ToString.Trim)
                        End If

                        Debug.WriteLine("strRemoteBuildParts2Split = " & strRemoteBuildParts2Split.ToString)

                        If strRemoteBuildParts(1) > buildVersion And strRemoteBuildParts2Split.Contains(buildVersion.ToString) = True Then
                            Dim updateQuestion As MsgBoxResult = MsgBox("There is an update available but it's classified as a minor update.  It's not a required update so if you do not want to update the program at this time, it is OK to keep using the version you have." & vbCrLf & vbCrLf & "Do you want to download the update?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, Me.Text)
                            If updateQuestion = MsgBoxResult.Yes Then
                                downloadAndDoUpdate()
                                Exit Sub
                            End If
                        ElseIf strRemoteBuildParts(1) > buildVersion And strRemoteBuildParts2Split.Contains(buildVersion.ToString) = False Then
                            downloadAndDoUpdate()
                            Exit Sub
                        ElseIf Short.Parse(strRemoteBuildParts(1)) = buildVersion Then
                            Debug.WriteLine(strRemoteBuildParts(1))
                            MsgBox("You already have the latest version.", MsgBoxStyle.Information, Me.Text)
                        End If

                        strRemoteBuildParts2Split.Clear()
                        strRemoteBuildParts2Split = Nothing
                    ElseIf Short.TryParse(strRemoteBuild, shortRemoteBuild) = True Then
                        If shortRemoteBuild < buildVersion Then
                            MsgBox("Somehow you have a version that is newer than is listed on the product web site, wierd.", MsgBoxStyle.Information, Me.Text)
                        ElseIf shortRemoteBuild = buildVersion Then
                            MsgBox("You already have the latest version.", MsgBoxStyle.Information, Me.Text)
                        ElseIf shortRemoteBuild > buildVersion Then
                            MsgBox(String.Format("There is an updated version of {0}.  The update will now download.", programName), MsgBoxStyle.Information, Me.Text & " Version Checker")
                            downloadAndDoUpdate(True)
                        End If
                    End If
                Else
                    btnCheckForUpdates.Enabled = True
                    MsgBox("There was an error checking for updates.", MsgBoxStyle.Information, Me.Text)
                End If
            Catch ex As Exception
                ' Ok, we crashed but who cares.  We give an error message.
                'MsgBox("Error while checking for new version.", MsgBoxStyle.Information, Me.Text)
            Finally
                btnCheckForUpdates.Enabled = True
            End Try
        End If
    End Sub

    Public Function checkForInternetConnection() As Boolean
        If My.Computer.Network.IsAvailable = True Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function SHA160(ByVal filename As String) As String
        Dim SHA1Engine As New Security.Cryptography.SHA1CryptoServiceProvider

        Dim FileStream As New IO.FileStream(filename, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read, 10 * 1048576, IO.FileOptions.SequentialScan)
        Dim Output As Byte() = SHA1Engine.ComputeHash(FileStream)
        FileStream.Close()
        FileStream.Dispose()

        Dim result As String = BitConverter.ToString(Output).ToLower().Replace("-", "").Trim
        SHA160 = result
    End Function

    Public Function verifyChecksum(urlOfChecksumFile As String, fileToVerify As String, boolGiveUserAnErrorMessage As Boolean) As Boolean
        Dim checksumFromWeb As String = Nothing

        Dim httpHelper As httpHelper = internetFunctions.createNewHTTPHelperObject()
        If httpHelper.getWebData(urlOfChecksumFile, checksumFromWeb, False) = False Then
            If boolGiveUserAnErrorMessage = True Then
                MsgBox("There was an error downloading the checksum verification file. Update process aborted.", MsgBoxStyle.Critical, "YAWA2 (Yet Another WinApp2.ini) Updater")
            End If

            Return False
        Else
            ' Checks to see if we have a valid SHA1 file.
            If Regex.IsMatch(checksumFromWeb, "([a-zA-Z0-9]{40})") = True Then
                checksumFromWeb = Regex.Match(checksumFromWeb, "([a-zA-Z0-9]{40})").Groups(1).Value().ToLower.Trim()

                If SHA160(fileToVerify) = checksumFromWeb Then
                    Return True
                Else
                    If boolGiveUserAnErrorMessage = True Then
                        MsgBox("There was an error in the download, checksums don't match. Update process aborted.", MsgBoxStyle.Critical, "YAWA2 (Yet Another WinApp2.ini) Updater")
                    End If

                    Return False
                End If
            Else
                If boolGiveUserAnErrorMessage = True Then
                    MsgBox("Invalid SHA1 file detected. Update process aborted.", MsgBoxStyle.Critical, "YAWA2 (Yet Another WinApp2.ini) Updater")
                End If

                Return False
            End If
        End If
    End Function

    Private Sub extractFileFromZIPFile(pathToZIPFile As String, fileToExtract As String, fileToWriteExtractedFileTo As String)
        Dim zipFileObject As New ICSharpCode.SharpZipLib.Zip.ZipFile(pathToZIPFile)
        Dim zipFileEntry As ICSharpCode.SharpZipLib.Zip.ZipEntry = zipFileObject.GetEntry(fileToExtract)

        If zipFileEntry IsNot Nothing Then
            Dim fileStream As New IO.FileStream(fileToWriteExtractedFileTo, IO.FileMode.Create)
            zipFileObject.GetInputStream(zipFileEntry).CopyTo(fileStream)
            fileStream.Close()
            fileStream.Dispose()
        End If

        zipFileObject.Close()
    End Sub

    Private Sub downloadAndDoUpdate(Optional ByVal outputText As Boolean = False)
        Dim fileInfo As New IO.FileInfo(Application.ExecutablePath)
        Dim newExecutableFilePath As String = fileInfo.Name & ".new.exe"

        Dim httpHelper As httpHelper = internetFunctions.createNewHTTPHelperObject()

        If httpHelper.downloadFile(programZipFileURL, zipFileName, False, False) = False Then
            MsgBox("There was an error while downloading required files.", MsgBoxStyle.Critical, Me.Text)
            Exit Sub
        End If

        If IO.File.Exists(zipFileName) = True Then
            If verifyChecksum(programZipFileSHA1URL, zipFileName, True) = False Then
                IO.File.Delete(zipFileName)
                Exit Sub
            End If
        End If

        fileInfo = Nothing

        If IO.File.Exists(zipFileName) Then
            extractFileFromZIPFile(zipFileName, programFileNameInZIP, newExecutableFilePath)
            IO.File.Delete(zipFileName)

            If boolWinXP = True Then
                Process.Start(updaterFileName, String.Format("--file={0}{1}{0}", Chr(34), Application.ExecutablePath))
            Else
                Dim startInfo As New ProcessStartInfo
                startInfo.FileName = newExecutableFilePath
                startInfo.Arguments = "-update"
                If canIWriteToTheCurrentDirectory() = False Then startInfo.Verb = "runas"
                Process.Start(startInfo)

                Process.GetCurrentProcess.Kill()
            End If
        End If

        Me.Close()
        Application.Exit()
    End Sub

    Private Sub btnAbout_Click(sender As Object, e As EventArgs) Handles btnAbout.Click
        Dim version() As String = Application.ProductVersion.Split(".".ToCharArray) ' Gets the program version
        Dim stringBuilder As New System.Text.StringBuilder

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
        If chkUseSSL.Checked = True Then
            programFunctions.saveSettingToINIFile("useSSL", "True")
            boolUseSSL = True
        Else
            programFunctions.saveSettingToINIFile("useSSL", "False")
            boolUseSSL = False
        End If
    End Sub
End Class
