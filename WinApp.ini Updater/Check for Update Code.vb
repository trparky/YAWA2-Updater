Imports System.Management
Imports System.Text.RegularExpressions
Imports System.Xml

Module Check_for_Update_Code
    Public Const programZipFileURL = "www.toms-world.org/download/YAWA2 Updater.zip"
    Public Const programZipFileSHA1URL = "www.toms-world.org/download/YAWA2 Updater.zip.sha1"

    Public Const programFileNameInZIP As String = "YAWA2 Updater.exe"

    Public versionInfo As String() = Application.ProductVersion.Split(".")
    Public shortMajor As Short = Short.Parse(versionInfo(versionPieces.major).Trim)
    Public shortMinor As Short = Short.Parse(versionInfo(versionPieces.minor).Trim)
    Public shortBuild As Short = Short.Parse(versionInfo(versionPieces.build).Trim)

    Public Const webSiteURL As String = "www.toms-world.org/blog/yawa2updater"

    Public Const programUpdateCheckerXMLFile As String = "www.toms-world.org/updates/yawa2_update.xml"
    Public Const programName As String = "YAWA2 Updater"

    Public versionStringWithoutBuild As String = String.Format("{0}.{1}", versionInfo(versionPieces.major), versionInfo(versionPieces.minor))

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

    Private Sub extractFileFromZIPFile(memoryStream As IO.MemoryStream, fileToExtract As String, fileToWriteExtractedFileTo As String)
        memoryStream.Position = 0

        Using zipFileObject As New IO.Compression.ZipArchive(memoryStream)
            Dim zipFileEntry As IO.Compression.ZipArchiveEntry = zipFileObject.GetEntry(fileToExtract)

            If zipFileEntry IsNot Nothing Then
                Using zipFileEntryIOStream As IO.Stream = zipFileEntry.Open()
                    Using fileStream As New IO.FileStream(fileToWriteExtractedFileTo, IO.FileMode.Create)
                        zipFileEntryIOStream.CopyTo(fileStream)
                    End Using
                End Using
            End If
        End Using
    End Sub

    Public Sub downloadAndDoUpdate()
        Dim memStream As New IO.MemoryStream()
        Dim fileInfo As New IO.FileInfo(Application.ExecutablePath)
        Dim newExecutableFilePath As String = fileInfo.Name & ".new.exe"

        Dim httpHelper As httpHelper = internetFunctions.createNewHTTPHelperObject()

        If Not httpHelper.downloadFile(programZipFileURL, memStream, False) Then
            MsgBox("There was an error while downloading required files.", MsgBoxStyle.Critical, programName)
            Exit Sub
        End If

        If Not verifyChecksum(programZipFileSHA1URL, memStream, True) Then Exit Sub

        extractFileFromZIPFile(memStream, programFileNameInZIP, newExecutableFilePath)

        Dim startInfo As New ProcessStartInfo With {.FileName = newExecutableFilePath, .Arguments = "-update"}
        If Not programFunctions.canIWriteToTheCurrentDirectory() Then startInfo.Verb = "runas"
        Process.Start(startInfo)

        Process.GetCurrentProcess.Kill()

        Application.Exit()
    End Sub

    Private Function SHA160(ByRef memStream As IO.MemoryStream) As String
        Using SHA1Engine As New Security.Cryptography.SHA1CryptoServiceProvider
            memStream.Position = 0
            Dim Output As Byte() = SHA1Engine.ComputeHash(memStream)
            memStream.Position = 0
            Return BitConverter.ToString(Output).ToLower().Replace("-", "").Trim
        End Using
    End Function

    Public Function verifyChecksum(urlOfChecksumFile As String, ByRef memStream As IO.MemoryStream, boolGiveUserAnErrorMessage As Boolean) As Boolean
        Dim checksumFromWeb As String = Nothing

        Dim httpHelper As httpHelper = internetFunctions.createNewHTTPHelperObject()

        If Not internetFunctions.createNewHTTPHelperObject().getWebData(urlOfChecksumFile, checksumFromWeb, False) Then
            If boolGiveUserAnErrorMessage Then MsgBox("There was an error downloading the checksum verification file. Update process aborted.", MsgBoxStyle.Critical, "YAWA2 (Yet Another WinApp2.ini) Updater")
            Return False
        Else
            ' Checks to see if we have a valid SHA1 file.
            If Regex.IsMatch(checksumFromWeb, "([a-zA-Z0-9]{40})") Then
                checksumFromWeb = Regex.Match(checksumFromWeb, "([a-zA-Z0-9]{40})").Groups(1).Value().ToLower.Trim()

                If SHA160(memStream).Equals(checksumFromWeb, StringComparison.OrdinalIgnoreCase) Then : Return True
                Else
                    If boolGiveUserAnErrorMessage Then MsgBox("There was an error in the download, checksums don't match. Update process aborted.", MsgBoxStyle.Critical, "YAWA2 (Yet Another WinApp2.ini) Updater")
                    Return False
                End If
            Else
                If boolGiveUserAnErrorMessage Then MsgBox("Invalid SHA1 file detected. Update process aborted.", MsgBoxStyle.Critical, "YAWA2 (Yet Another WinApp2.ini) Updater")
                Return False
            End If
        End If
    End Function

    Public Sub checkForUpdates(parentForm As Form1)
        Dim xmlData As String = Nothing

        If internetFunctions.createNewHTTPHelperObject().getWebData(programUpdateCheckerXMLFile, xmlData, False) Then
            Dim remoteVersion As String = Nothing
            Dim remoteBuild As String = Nothing
            Dim response As processUpdateXMLResponse = processUpdateXMLData(xmlData, remoteVersion, remoteBuild)

            If response = processUpdateXMLResponse.newVersion Then
                downloadAndDoUpdate()
            ElseIf response = processUpdateXMLResponse.parseError Or response = processUpdateXMLResponse.exceptionError Then
                MsgBox("There was an error when trying to parse response from server.", MsgBoxStyle.Critical, programName)
            ElseIf response = processUpdateXMLResponse.newerVersionThanWebSite Then
                MsgBox("This is weird, you have a version that's newer than what's listed on the web site.", MsgBoxStyle.Information, programName)
            ElseIf response = processUpdateXMLResponse.noUpdateNeeded Then
                MsgBox("You already have the latest version.", MsgBoxStyle.Information, programName)
            End If
        Else
            parentForm.Invoke(Sub() parentForm.btnCheckForUpdates.Enabled = True)
            MsgBox("There was an error checking for updates.", MsgBoxStyle.Information, programName)
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

    Public Sub searchForProcessAndKillIt(strFileName As String, boolFullFilePathPassed As Boolean)
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

    Public Sub newFileDeleter()
        If IO.File.Exists(Application.ExecutablePath & ".new.exe") = True Then
            Threading.ThreadPool.QueueUserWorkItem(Sub()
                                                       searchForProcessAndKillIt(Application.ExecutablePath & ".new.exe", False)
                                                       IO.File.Delete(Application.ExecutablePath & ".new.exe")
                                                   End Sub)
        End If
    End Sub
End Module

Public Enum versionPieces As Short
    major = 0
    minor = 1
    build = 2
    revision = 3
End Enum