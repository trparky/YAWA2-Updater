Imports System.Xml.Serialization
Imports System.IO

Namespace AppSettings
    Public Structure AppSettings
        Public boolMobileMode, boolTrim, boolNotifyAfterUpdateAtLogon, boolSleepOnSilentStartup As Boolean
        Public strCustomEntries As String, shortSleepOnSilentStartup As Short
    End Structure

    Module Globals
        Public AppSettingsObject As AppSettings

        ' All operations on the settings file are done using atomic file transactions. This means that until we have verified that
        ' all data has been written to disk after changing the settings file the original file remains intact because the original
        ' settings file isn't the actual file that has been operated on; only a temporary file has been worked on.
        '
        ' These are the steps that the program executes when updating or writing the settings file.
        ' 1. Acquire a mutex lock. If another part of the program or another instance of the program has a mutex lock then
        '    we wait until the mutex lock has been released and then acquire it after the lock has been released.
        ' 2. Serialize the settings object and write the new data to a MemoryStream.
        ' 3. Write all data that's in the MemoryStream to a temporary file on disk.
        ' 4. Verify that all data has been written to disk by comparing what's in the MemoryStream to what's been written to disk.
        ' 5. Once we have verified that all data has been successfully written to disk we then delete the original settings file
        '    and rename the temporary file to same name of the original file.
        ' 6. And now our atomic file transaction is complete.
        ' 7. Release the mutex lock.
        '
        ' Yes, this process requires more code and time to complete but it ensures that the integrity of the settings file is maintained
        ' at all times and as much as humanly possible. All of this is to ensure that if the program crashes during a settings file
        ' write operation the settings file will maintain its integrity.

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
                Dim boolSuccessfulWriteToDisk As Boolean = False ' Declare a variable to store our results of the write operation.

                ' Create a MemoryStream to temporarily write our serialized data to.
                Using memoryStream As New MemoryStream()
                    ' Create an XMLSerialized to write our serialized data to our MemoryStream.
                    Dim xmlSerializerObject As New XmlSerializer(AppSettingsObject.GetType)

                    ' And serialize it.
                    xmlSerializerObject.Serialize(memoryStream, AppSettingsObject)

                    ' Write all data in our MemoryStream to disk.
                    File.WriteAllBytes(programConstants.configXMLFile & ".temp", memoryStream.ToArray())

                    ' This validates the data that has been written to disk in the form of a temporary file.
                    boolSuccessfulWriteToDisk = VerifyDataOnDisk(memoryStream, programConstants.configXMLFile & ".temp")
                End Using

                If boolSuccessfulWriteToDisk Then
                    ' Delete the current log file.
                    DeleteFileWithNoException(programConstants.configXMLFile)

                    ' And now our atomic file transaction is complete. The data that has been written to disk
                    ' has been validated so we can move the new log file into place of the old log file.
                    File.Move(programConstants.configXMLFile & ".temp", programConstants.configXMLFile)
                Else
                    DeleteFileWithNoException(programConstants.configXMLFile & ".temp")
                    MsgBox("There was an error while writing the settings file to disk, the original settings file has been preserved to prevent data corruption.", MsgBoxStyle.Critical, "WinApp.ini Updater")
                End If
            End SyncLock
        End Sub

        ''' <summary>Writes the contents of a MemoryStream to disk and verifies the contents of both. It returns a Boolean value indicating the results of the file operation.</summary>
        ''' <param name="memoryStream">The MemoryStream Object that you want to write to disk.</param>
        ''' <param name="strPathOfFileToVerify">The path of the file to which you want to write the contents of the MemoryStream to.</param>
        ''' <returns>A Boolean value indicating the results of the file operation.</returns>
        Private Function VerifyDataOnDisk(ByRef memoryStream As MemoryStream, strPathOfFileToVerify As String) As Boolean
            Dim tempFileByteArray As Byte() = File.ReadAllBytes(strPathOfFileToVerify)
            Return tempFileByteArray.Length = memoryStream.Length AndAlso tempFileByteArray.SequenceEqual(memoryStream.ToArray)
        End Function

        ''' <summary>Delete a file with no Exception. It first checks to see if the file exists and if it does it attempts to delete it. If it fails to do so then it will fail silently without an Exception.</summary>
        Public Sub DeleteFileWithNoException(pathToFile As String)
            Try
                If File.Exists(pathToFile) Then File.Delete(pathToFile)
            Catch ex As Exception
            End Try
        End Sub
    End Module
End Namespace