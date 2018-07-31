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

    ''' <summary>This parses the XML updata data and determines if an update is needed.</summary>
    ''' <param name="xmlData">The XML data from the web site.</param>
    ''' <returns>A Boolean value indicating if the program has been updated or not.</returns>
    Public Function processUpdateXMLData(ByVal xmlData As String) As Boolean
        Try
            Dim xmlDocument As New XmlDocument() ' First we create an XML Document Object.
            xmlDocument.Load(New IO.StringReader(xmlData)) ' Now we try and parse the XML data.

            Dim xmlNode As XmlNode = xmlDocument.SelectSingleNode("/xmlroot")

            Dim remoteVersion As String = xmlNode.SelectSingleNode("version").InnerText.Trim
            Dim remoteBuild As String = xmlNode.SelectSingleNode("build").InnerText.Trim
            Dim shortRemoteBuild As Short

            ' This checks to see if current version and the current build matches that of the remote values in the XML document.
            If remoteVersion.Equals(versionStringWithoutBuild) And remoteBuild.Equals(shortBuild.ToString) Then
                If Short.TryParse(remoteBuild, shortRemoteBuild) And remoteVersion.Equals(versionStringWithoutBuild) Then
                    If shortRemoteBuild < shortBuild Then
                        ' This is weird, the remote build is less than the current build. Something went wrong. So to be safe we're going to return a False value indicating that there is no update to download. Better to be safe.
                        Return False
                    End If
                End If

                ' OK, they match so there's no update to download and update to therefore we return a False value.
                Return False
            Else
                ' We return a True value indicating that there is a new version to download and install.
                Return True
            End If
        Catch ex As XPath.XPathException
            ' Something went wrong so we return a False value.
            Return False
        Catch ex As XmlException
            ' Something went wrong so we return a False value.
            Return False
        Catch ex As Exception
            ' Something went wrong so we return a False value.
            Return False
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

    Public Sub downloadAndDoUpdate(Optional ByVal outputText As Boolean = False)
        Dim memStream As New IO.MemoryStream()
        Dim fileInfo As New IO.FileInfo(Application.ExecutablePath)
        Dim newExecutableFilePath As String = fileInfo.Name & ".new.exe"

        Dim httpHelper As httpHelper = internetFunctions.createNewHTTPHelperObject()

        If Not httpHelper.downloadFile(programZipFileURL, memStream, False) Then
            MsgBox("There was an error while downloading required files.", MsgBoxStyle.Critical, programName)
            Exit Sub
        End If

        If Not verifyChecksum(programZipFileSHA1URL, memStream, True) Then Exit Sub

        fileInfo = Nothing

        extractFileFromZIPFile(memStream, programFileNameInZIP, newExecutableFilePath)

        If boolWinXP Then : Process.Start(newExecutableFilePath, "-update")
        Else
            Dim startInfo As New ProcessStartInfo With {
                .FileName = newExecutableFilePath,
                .Arguments = "-update"
            }
            If Not canIWriteToTheCurrentDirectory() Then startInfo.Verb = "runas"
            Process.Start(startInfo)

            Process.GetCurrentProcess.Kill()
        End If

        Application.Exit()
    End Sub

    Private Function SHA160(ByRef memStream As IO.MemoryStream) As String
        Dim SHA1Engine As New Security.Cryptography.SHA1CryptoServiceProvider
        memStream.Position = 0
        Dim Output As Byte() = SHA1Engine.ComputeHash(memStream)
        memStream.Position = 0
        Return BitConverter.ToString(Output).ToLower().Replace("-", "").Trim
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
            If processUpdateXMLData(xmlData) Then : downloadAndDoUpdate()
            Else : MsgBox("You already have the latest version.", MsgBoxStyle.Information, programName)
            End If
        Else
            parentForm.Invoke(Sub() parentForm.btnCheckForUpdates.Enabled = True)
            MsgBox("There was an error checking for updates.", MsgBoxStyle.Information, programName)
        End If
    End Sub

    Private Function doesPIDExist(PID As Integer) As Boolean
        Using searcher As New ManagementObjectSearcher("root\CIMV2", String.Format("SELECT * FROM Win32_Process WHERE ProcessId={0}", PID))
            Return If(searcher.Get.Count = 0, False, True)
        End Using
    End Function

    Private Sub killProcess(PID As Integer)
        Dim processDetail As Process

        processDetail = Process.GetProcessById(PID)
        processDetail.Kill()

        Threading.Thread.Sleep(100)

        If doesPIDExist(PID) Then killProcess(PID)
    End Sub

    Public Sub searchForProcessAndKillIt(fileName As String)
        Dim fullFileName As String = New IO.FileInfo(fileName).FullName
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

    Public Sub newFileDeleter()
        If IO.File.Exists(Application.ExecutablePath & ".new.exe") = True Then
            Threading.ThreadPool.QueueUserWorkItem(Sub()
                                                       searchForProcessAndKillIt(Application.ExecutablePath & ".new.exe")
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