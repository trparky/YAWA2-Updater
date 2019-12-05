Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Xml
Imports System.Security.AccessControl
Imports System.Security.Principal

Module checkForUpdateModules
    Public Const strMessageBoxTitleText As String = "YAWA2 (Yet Another WinApp2.ini) Updater"
    Public Const strProgramName As String = "YAWA2 (Yet Another WinApp2.ini) Updater"
    Private Const strZipFileName As String = "YAWA2 Updater.zip"

    Public Sub doUpdateAtStartup()
        If File.Exists(strZipFileName) Then File.Delete(strZipFileName)
        Dim currentProcessFileName As String = New FileInfo(Application.ExecutablePath).Name

        If currentProcessFileName.caseInsensitiveContains(".new.exe") Then
            Dim mainEXEName As String = currentProcessFileName.caseInsensitiveReplace(".new.exe", "")
            searchForProcessAndKillIt(mainEXEName, False)

            File.Delete(mainEXEName)
            File.Copy(currentProcessFileName, mainEXEName)

            Process.Start(New ProcessStartInfo With {.FileName = mainEXEName})
            Process.GetCurrentProcess.Kill()
        Else
            MsgBox("The environment is not ready for an update. This process will now terminate.", MsgBoxStyle.Critical, strMessageBoxTitleText)
            Process.GetCurrentProcess.Kill()
        End If
    End Sub

    ''' <summary>Checks to see if a Process ID or PID exists on the system.</summary>
    ''' <param name="PID">The PID of the process you are checking the existance of.</param>
    ''' <param name="processObject">If the PID does exist, the function writes back to this argument in a ByRef way a Process Object that can be interacted with outside of this function.</param>
    ''' <returns>Return a Boolean value. If the PID exists, it return a True value. If the PID doesn't exist, it returns a False value.</returns>
    Private Function doesProcessIDExist(ByVal PID As Integer, ByRef processObject As Process) As Boolean
        Try
            processObject = Process.GetProcessById(PID)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Sub newFileDeleter()
        If IO.File.Exists(Application.ExecutablePath & ".new.exe") Then
            Threading.ThreadPool.QueueUserWorkItem(Sub()
                                                       searchForProcessAndKillIt(New FileInfo(Application.ExecutablePath).Name & ".new.exe", False)
                                                       File.Delete(Application.ExecutablePath & ".new.exe")
                                                   End Sub)
        End If
    End Sub

    Private Sub killProcess(processID As Integer)
        Dim processObject As Process = Nothing

        ' First we are going to check if the Process ID exists.
        If doesProcessIDExist(processID, processObject) Then
            Try
                processObject.Kill() ' Yes, it does so let's kill it.
            Catch ex As Exception
                ' Wow, it seems that even with double-checking if a process exists by it's PID number things can still go wrong.
                ' So this Try-Catch block is here to trap any possible errors when trying to kill a process by it's PID number.
            End Try
        End If

        Threading.Thread.Sleep(250) ' We're going to sleep to give the system some time to kill the process.

        '' Now we are going to check again if the Process ID exists and if it does, we're going to attempt to kill it again.
        If doesProcessIDExist(processID, processObject) Then
            Try
                processObject.Kill()
            Catch ex As Exception
                ' Wow, it seems that even with double-checking if a process exists by it's PID number things can still go wrong.
                ' So this Try-Catch block is here to trap any possible errors when trying to kill a process by it's PID number.
            End Try
        End If

        Threading.Thread.Sleep(250) ' We're going to sleep (again) to give the system some time to kill the process.
    End Sub

    Private Function getProcessExecutablePath(processID As Integer) As String
        Dim memoryBuffer As New Text.StringBuilder(1024)
        Dim processHandle As IntPtr = NativeMethod.NativeMethods.OpenProcess(NativeMethod.ProcessAccessFlags.PROCESS_QUERY_LIMITED_INFORMATION, False, processID)

        If processHandle <> IntPtr.Zero Then
            Try
                Dim memoryBufferSize As Integer = memoryBuffer.Capacity

                If NativeMethod.NativeMethods.QueryFullProcessImageName(processHandle, 0, memoryBuffer, memoryBufferSize) Then
                    Return memoryBuffer.ToString()
                End If
            Finally
                NativeMethod.NativeMethods.CloseHandle(processHandle)
            End Try
        End If

        NativeMethod.NativeMethods.CloseHandle(processHandle)
        Return Nothing
    End Function

    Private Sub searchForProcessAndKillIt(strFileName As String, boolFullFilePathPassed As Boolean)
        Dim processExecutablePath As String
        Dim processExecutablePathFileInfo As IO.FileInfo

        For Each process As Process In Process.GetProcesses()
            processExecutablePath = getProcessExecutablePath(process.Id)

            If processExecutablePath IsNot Nothing Then
                Try
                    processExecutablePathFileInfo = New IO.FileInfo(processExecutablePath)

                    If boolFullFilePathPassed Then
                        If strFileName.Equals(processExecutablePathFileInfo.FullName, StringComparison.OrdinalIgnoreCase) Then
                            killProcess(process.Id)
                        End If
                    Else
                        If strFileName.Equals(processExecutablePathFileInfo.Name, StringComparison.OrdinalIgnoreCase) Then
                            killProcess(process.Id)
                        End If
                    End If
                Catch ex As ArgumentException
                End Try
            End If
        Next
    End Sub
End Module

Class Check_for_Update_Stuff
    Private Const programZipFileURL = "www.toms-world.org/download/YAWA2 Updater.zip"
    Private Const programZipFileSHA256URL = "www.toms-world.org/download/YAWA2 Updater.zip.sha2"
    Private Const programFileNameInZIP As String = "YAWA2 Updater.exe"
    Private Const programUpdateCheckerXMLFile As String = "www.toms-world.org/updates/yawa2_update.xml"

    Public windowObject As Form1
    Public Shared versionInfo As String() = Application.ProductVersion.Split(".")
    Private shortBuild As Short = Short.Parse(versionInfo(versionPieces.build).Trim)
    Public Shared versionString As String = String.Format("{0}.{1} Build {2}", versionInfo(0), versionInfo(1), versionInfo(2))
    Private versionStringWithoutBuild As String = String.Format("{0}.{1}", versionInfo(versionPieces.major), versionInfo(versionPieces.minor))

    Public Sub New(inputWindowObject As Form1)
        windowObject = inputWindowObject
    End Sub

    Private Shared Sub extractFileFromZIPFile(ByRef memoryStream As MemoryStream, fileToExtract As String, fileToWriteExtractedFileTo As String)
        Using zipFileObject As New Compression.ZipArchive(memoryStream, Compression.ZipArchiveMode.Read)
            Using fileStream As New FileStream(fileToWriteExtractedFileTo, FileMode.Create)
                zipFileObject.GetEntry(fileToExtract).Open().CopyTo(fileStream)
            End Using
        End Using
    End Sub

    Enum processUpdateXMLResponse As Short
        noUpdateNeeded
        newVersion
        newerVersionThanWebSite
        parseError
        exceptionError
    End Enum

    ''' <summary>This parses the XML updata data and determines if an update is needed.</summary>
    ''' <param name="xmlData">The XML data from the web site.</param>
    ''' <returns>A Boolean value indicating if the program has been updated or not.</returns>
    Public Function processUpdateXMLData(ByVal xmlData As String, ByRef remoteVersion As String, ByRef remoteBuild As String) As processUpdateXMLResponse
        Try
            Dim xmlDocument As New XmlDocument() ' First we create an XML Document Object.
            xmlDocument.Load(New IO.StringReader(xmlData)) ' Now we try and parse the XML data.

            Dim xmlNode As XmlNode = xmlDocument.SelectSingleNode("/xmlroot")

            remoteVersion = xmlNode.SelectSingleNode("version").InnerText.Trim
            remoteBuild = xmlNode.SelectSingleNode("build").InnerText.Trim
            Dim shortRemoteBuild As Short

            ' This checks to see if current version and the current build matches that of the remote values in the XML document.
            If remoteVersion.Equals(versionStringWithoutBuild) And remoteBuild.Equals(shortBuild.ToString) Then
                ' Both the remoteVersion and the remoteBuild equals that of the current version,
                ' therefore we return a sameVersion value indicating no update is required.
                Return processUpdateXMLResponse.noUpdateNeeded
            Else
                ' First we do a check of the version, if it's not equal we simply return a newVersion value.
                If Not remoteVersion.Equals(versionStringWithoutBuild) Then
                    ' We return a newVersion value indicating that there is a new version to download and install.
                    Return processUpdateXMLResponse.newVersion
                Else
                    ' Now let's do some sanity checks here. 
                    If Short.TryParse(remoteBuild, shortRemoteBuild) Then
                        If shortRemoteBuild < shortBuild Then
                            ' This is weird, the remote build is less than the current build so we return a newerVersionThanWebSite value.
                            Return processUpdateXMLResponse.newerVersionThanWebSite
                        ElseIf shortRemoteBuild.Equals(shortBuild) Then
                            ' The build numbers match, therefore therefore we return a sameVersion value.
                            Return processUpdateXMLResponse.noUpdateNeeded
                        End If
                    Else
                        ' Something went wrong, we couldn't parse the value of the remoteBuild number so we return a parseError value.
                        Return processUpdateXMLResponse.parseError
                    End If

                    ' We return a newVersion value indicating that there is a new version to download and install.
                    Return processUpdateXMLResponse.newVersion
                End If
            End If
        Catch ex As Exception
            ' Something went wrong so we return a exceptionError value.
            Return processUpdateXMLResponse.exceptionError
        End Try
    End Function

    Private Shared Function canIWriteToTheCurrentDirectory() As Boolean
        Return canIWriteThere(New FileInfo(Application.ExecutablePath).DirectoryName)
    End Function

    Private Shared Function randomString(length As Integer) As String
        Dim random As Random = New Random()
        Dim builder As New Text.StringBuilder()
        Dim ch As Char
        Dim legalCharacters As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890"

        For cntr As Integer = 0 To length
            ch = legalCharacters.Substring(random.Next(0, legalCharacters.Length), 1)
            builder.Append(ch)
        Next

        Return builder.ToString()
    End Function

    Private Shared Function canIWriteThere(folderPath As String) As Boolean
        ' We make sure we get valid folder path by taking off the leading slash.
        If folderPath.EndsWith("\") Then folderPath = folderPath.Substring(0, folderPath.Length - 1)

        If String.IsNullOrEmpty(folderPath) Or Not Directory.Exists(folderPath) Then Return False

        If checkByFolderACLs(folderPath) Then
            Try
                Dim strRandomFileName As String = randomString(15) & ".txt"
                File.Create(Path.Combine(folderPath, strRandomFileName), 1, FileOptions.DeleteOnClose).Close()
                If File.Exists(Path.Combine(folderPath, strRandomFileName)) Then File.Delete(Path.Combine(folderPath, strRandomFileName))
                Return True
            Catch ex As Exception
                Return False
            End Try
        Else
            Return False
        End If
    End Function

    Private Shared Function checkByFolderACLs(folderPath As String) As Boolean
        Try
            Dim directoryACLs As DirectorySecurity = Directory.GetAccessControl(folderPath)
            Dim directoryUsers As String = WindowsIdentity.GetCurrent.User.Value
            Dim directoryAccessRights As FileSystemAccessRule
            Dim fileSystemRightsVariable As FileSystemRights

            For Each rule As AuthorizationRule In directoryACLs.GetAccessRules(True, True, GetType(SecurityIdentifier))
                If rule.IdentityReference.Value = directoryUsers Then
                    directoryAccessRights = DirectCast(rule, FileSystemAccessRule)

                    If directoryAccessRights.AccessControlType = AccessControlType.Allow Then
                        fileSystemRightsVariable = directoryAccessRights.FileSystemRights

                        If fileSystemRightsVariable = (FileSystemRights.Read Or FileSystemRights.Modify Or FileSystemRights.Write Or FileSystemRights.FullControl) Then
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

    Private Shared Function createNewHTTPHelperObject() As httpHelper
        Dim httpHelper As New httpHelper With {
            .setUserAgent = createHTTPUserAgentHeaderString(),
            .useHTTPCompression = True,
            .setProxyMode = True
        }
        httpHelper.addHTTPHeader("PROGRAM_NAME", strProgramName)
        httpHelper.addHTTPHeader("PROGRAM_VERSION", versionString)
        httpHelper.addHTTPHeader("OPERATING_SYSTEM", getFullOSVersionString())

        httpHelper.setURLPreProcessor = Function(ByVal strURLInput As String) As String
                                            Try
                                                If Not strURLInput.Trim.ToLower.StartsWith("http") Then
                                                    Return If(programFunctions.loadSettingFromINIFile("useSSL", True), "https://", "http://") & strURLInput
                                                Else
                                                    Return strURLInput
                                                End If
                                            Catch ex As Exception
                                                Return strURLInput
                                            End Try
                                        End Function

        Return httpHelper
    End Function

    Private Shared Function SHA256ChecksumStream(ByRef stream As Stream) As String
        Using SHA256Engine As New Security.Cryptography.SHA256CryptoServiceProvider
            Return BitConverter.ToString(SHA256Engine.ComputeHash(stream)).ToLower().Replace("-", "").Trim
        End Using
    End Function

    Private Function verifyChecksum(urlOfChecksumFile As String, ByRef memStream As MemoryStream, ByRef httpHelper As httpHelper, boolGiveUserAnErrorMessage As Boolean) As Boolean
        Dim checksumFromWeb As String = Nothing
        memStream.Position = 0

        Try
            If httpHelper.getWebData(urlOfChecksumFile, checksumFromWeb) Then
                Dim regexObject As New Regex("([a-zA-Z0-9]{64})")

                ' Checks to see if we have a valid SHA256 file.
                If regexObject.IsMatch(checksumFromWeb) Then
                    ' Now that we have a valid SHA256 file we need to parse out what we want.
                    checksumFromWeb = regexObject.Match(checksumFromWeb).Groups(1).Value.Trim()

                    ' Now we do the actual checksum verification by passing the name of the file to the SHA256() function
                    ' which calculates the checksum of the file on disk. We then compare it to the checksum from the web.
                    If SHA256ChecksumStream(memStream).Equals(checksumFromWeb, StringComparison.OrdinalIgnoreCase) Then
                        Return True ' OK, things are good; the file passed checksum verification so we return True.
                    Else
                        ' The checksums don't match. Oops.
                        If boolGiveUserAnErrorMessage Then
                            windowObject.Invoke(Sub() MsgBox("There was an error in the download, checksums don't match. Update process aborted.", MsgBoxStyle.Critical, strMessageBoxTitleText))
                        End If

                        Return False
                    End If
                Else
                    If boolGiveUserAnErrorMessage Then
                        windowObject.Invoke(Sub() MsgBox("Invalid SHA2 file detected. Update process aborted.", MsgBoxStyle.Critical, strMessageBoxTitleText))
                    End If

                    Return False
                End If
            Else
                If boolGiveUserAnErrorMessage Then
                    windowObject.Invoke(Sub() MsgBox("There was an error downloading the checksum verification file. Update process aborted.", MsgBoxStyle.Critical, strMessageBoxTitleText))
                End If

                Return False
            End If
        Catch ex As Exception
            If boolGiveUserAnErrorMessage Then
                windowObject.Invoke(Sub() MsgBox("There was an error downloading the checksum verification file. Update process aborted.", MsgBoxStyle.Critical, strMessageBoxTitleText))
            End If

            Return False
        End Try
    End Function

    Private Sub downloadAndPerformUpdate()
        Dim fileInfo As New FileInfo(Application.ExecutablePath)
        Dim newExecutableName As String = fileInfo.Name & ".new.exe"

        Dim httpHelper As httpHelper = createNewHTTPHelperObject()

        Using memoryStream As New MemoryStream()
            If Not httpHelper.downloadFile(programZipFileURL, memoryStream, False) Then
                windowObject.Invoke(Sub() MsgBox("There was an error while downloading required files.", MsgBoxStyle.Critical, strMessageBoxTitleText))
                Exit Sub
            End If

            If Not verifyChecksum(programZipFileSHA256URL, memoryStream, httpHelper, True) Then
                windowObject.Invoke(Sub() MsgBox("There was an error while downloading required files.", MsgBoxStyle.Critical, strMessageBoxTitleText))
                Exit Sub
            End If

            fileInfo = Nothing
            memoryStream.Position = 0

            extractFileFromZIPFile(memoryStream, programFileNameInZIP, newExecutableName)
        End Using

        Dim startInfo As New ProcessStartInfo With {
            .FileName = newExecutableName,
            .Arguments = "-update"
        }
        If Not canIWriteToTheCurrentDirectory() Then startInfo.Verb = "runas"
        Process.Start(startInfo)

        Process.GetCurrentProcess.Kill()
    End Sub

    ''' <summary>Creates a User Agent String for this program to be used in HTTP requests.</summary>
    ''' <returns>String type.</returns>
    Private Shared Function createHTTPUserAgentHeaderString() As String
        Dim versionInfo As String() = Application.ProductVersion.Split(".")
        Dim versionString As String = String.Format("{0}.{1} Build {2}", versionInfo(0), versionInfo(1), versionInfo(2))
        Return String.Format("{2} version {0} on {1}", versionString, getFullOSVersionString(), strProgramName)
    End Function

    Private Shared Function getFullOSVersionString() As String
        Try
            Dim intOSMajorVersion As Integer = Environment.OSVersion.Version.Major
            Dim intOSMinorVersion As Integer = Environment.OSVersion.Version.Minor
            Dim dblDOTNETVersion As Double = Double.Parse(Environment.Version.Major & "." & Environment.Version.Minor)
            Dim strOSName As String

            If intOSMajorVersion = 5 And intOSMinorVersion = 0 Then
                strOSName = "Windows 2000"
            ElseIf intOSMajorVersion = 5 And intOSMinorVersion = 1 Then
                strOSName = "Windows XP"
            ElseIf intOSMajorVersion = 6 And intOSMinorVersion = 0 Then
                strOSName = "Windows Vista"
            ElseIf intOSMajorVersion = 6 And intOSMinorVersion = 1 Then
                strOSName = "Windows 7"
            ElseIf intOSMajorVersion = 6 And intOSMinorVersion = 2 Then
                strOSName = "Windows 8"
            ElseIf intOSMajorVersion = 6 And intOSMinorVersion = 3 Then
                strOSName = "Windows 8.1"
            ElseIf intOSMajorVersion = 10 Then
                strOSName = "Windows 10"
            Else
                strOSName = String.Format("Windows NT {0}.{1}", intOSMajorVersion, intOSMinorVersion)
            End If

            Return String.Format("{0} {2}-bit (Microsoft .NET {1})", strOSName, dblDOTNETVersion, If(Environment.Is64BitOperatingSystem, "64", "32"))
        Catch ex As Exception
            Try
                Return "Unknown Windows Operating System (" & Environment.OSVersion.VersionString & ")"
            Catch ex2 As Exception
                Return "Unknown Windows Operating System"
            End Try
        End Try
    End Function

    Private Function BackgroundThreadMessageBox(ByVal strMsgBoxPrompt As String, ByVal style As MsgBoxStyle, ByVal strMsgBoxTitle As String) As MsgBoxResult
        If windowObject.InvokeRequired Then
            Return CType(windowObject.Invoke(New Func(Of MsgBoxResult)(Function() MsgBox(strMsgBoxPrompt, style, strMsgBoxTitle))), MsgBoxResult)
        Else
            Return MsgBox(strMsgBoxPrompt, style, strMsgBoxTitle)
        End If
    End Function

    Public Sub checkForUpdates(Optional boolShowMessageBox As Boolean = True)
        windowObject.Invoke(Sub()
                                windowObject.btnCheckForUpdates.Enabled = False
                            End Sub)

        If Not My.Computer.Network.IsAvailable Then
            windowObject.Invoke(Sub() MsgBox("No Internet connection detected.", MsgBoxStyle.Information, strMessageBoxTitleText))
        Else
            Try
                Dim xmlData As String = Nothing
                Dim httpHelper As httpHelper = createNewHTTPHelperObject()

                If httpHelper.getWebData(programUpdateCheckerXMLFile, xmlData, False) Then
                    Dim remoteVersion As String = Nothing
                    Dim remoteBuild As String = Nothing
                    Dim response As processUpdateXMLResponse = processUpdateXMLData(xmlData, remoteVersion, remoteBuild)

                    If response = processUpdateXMLResponse.newVersion Then
                        If BackgroundThreadMessageBox(String.Format("An update to {2} (version {0} Build {1}) is available to be downloaded, do you want to download and update to this new version?", remoteVersion, remoteBuild, strProgramName), MsgBoxStyle.Question + MsgBoxStyle.YesNo, strMessageBoxTitleText) = MsgBoxResult.Yes Then
                            downloadAndPerformUpdate()
                        Else
                            windowObject.Invoke(Sub() MsgBox("The update will not be downloaded.", MsgBoxStyle.Information, strMessageBoxTitleText))
                        End If
                    ElseIf response = processUpdateXMLResponse.noUpdateNeeded Then
                        If boolShowMessageBox Then windowObject.Invoke(Sub() MsgBox("You already have the latest version, there is no need to update this program.", MsgBoxStyle.Information, strMessageBoxTitleText))
                    ElseIf response = processUpdateXMLResponse.parseError Or response = processUpdateXMLResponse.exceptionError Then
                        If boolShowMessageBox Then windowObject.Invoke(Sub() MsgBox("There was an error when trying to parse the response from the server.", MsgBoxStyle.Critical, strMessageBoxTitleText))
                    ElseIf response = processUpdateXMLResponse.newerVersionThanWebSite Then
                        If boolShowMessageBox Then windowObject.Invoke(Sub() MsgBox("This is weird, you have a version that's newer than what's listed on the web site.", MsgBoxStyle.Information, strMessageBoxTitleText))
                    End If
                Else
                    If boolShowMessageBox Then windowObject.Invoke(Sub() MsgBox("There was an error checking for updates.", MsgBoxStyle.Information, strMessageBoxTitleText))
                End If
            Catch ex As Exception
                ' Ok, we crashed but who cares.
            Finally
                windowObject.Invoke(Sub()
                                        windowObject.btnCheckForUpdates.Enabled = True
                                    End Sub)
                windowObject = Nothing
            End Try
        End If
    End Sub
End Class

Public Enum versionPieces As Short
    major = 0
    minor = 1
    build = 2
    revision = 3
End Enum