Imports System.Xml.Serialization
Imports System.IO

Namespace AppSettings
    Public Structure AppSettings
        Public boolMobileMode, boolTrim, boolNotifyAfterUpdateAtLogon, boolSleepOnSilentStartup As Boolean
        Public strCustomEntries As String, shortSleepOnSilentStartup As Short
    End Structure

    Module Globals
        Public AppSettingsObject As AppSettings

        Public Sub LoadSettingsFromXMLFileAppSettings()
            SyncLock programFunctions.LockObject
                AppSettingsObject = New AppSettings

                Using streamReader As New StreamReader(programConstants.configXMLFile)
                    Dim xmlSerializerObject As New XmlSerializer(AppSettingsObject.GetType)
                    AppSettingsObject = xmlSerializerObject.Deserialize(streamReader)
                End Using
            End SyncLock
        End Sub

        Public Sub SaveSettingsToXMLFile()
            SyncLock programFunctions.LockObject
                Dim boolSuccessfulWriteToDisk As Boolean = False

                Using memoryStream As New MemoryStream()
                    Dim xmlSerializerObject As New XmlSerializer(AppSettingsObject.GetType)
                    xmlSerializerObject.Serialize(memoryStream, AppSettingsObject)
                    File.WriteAllBytes(programConstants.configXMLFile & ".temp", memoryStream.ToArray())
                    boolSuccessfulWriteToDisk = VerifyDataOnDisk(memoryStream, programConstants.configXMLFile & ".temp")
                End Using

                If boolSuccessfulWriteToDisk Then
                    DeleteFileWithNoException(programConstants.configXMLFile)
                    File.Move(programConstants.configXMLFile & ".temp", programConstants.configXMLFile)
                Else
                    DeleteFileWithNoException(programConstants.configXMLFile & ".temp")
                    MsgBox("There was an error while writing the settings file to disk, the original settings file has been preserved to prevent data corruption.", MsgBoxStyle.Critical, "WinApp.ini Updater")
                End If
            End SyncLock
        End Sub

        Private Function VerifyDataOnDisk(ByRef memoryStream As MemoryStream, strPathOfFileToVerify As String) As Boolean
            Dim tempFileByteArray As Byte() = File.ReadAllBytes(strPathOfFileToVerify)
            Return tempFileByteArray.Length = memoryStream.Length AndAlso tempFileByteArray.SequenceEqual(memoryStream.ToArray)
        End Function

        Public Sub DeleteFileWithNoException(pathToFile As String)
            Try
                If File.Exists(pathToFile) Then File.Delete(pathToFile)
            Catch ex As Exception
            End Try
        End Sub
    End Module
End Namespace