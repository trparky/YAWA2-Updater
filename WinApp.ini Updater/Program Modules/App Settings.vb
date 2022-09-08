Imports System.Xml.Serialization
Imports YAWA2_Updater

Namespace Global
    Public Structure AppSettings
        Public boolMobileMode, boolTrim, boolNotifyAfterUpdateAtLogon, boolSleepOnSilentStartup As Boolean
        Public strCustomEntries As String, shortSleepOnSilentStartup As Short
    End Structure

    Module Globals
        Public AppSettingsObject As AppSettings

        Public Sub LoadSettingsFromXMLFileAppSettings()
            SyncLock programFunctions.LockObject
                AppSettingsObject = New AppSettings

                Using streamReader As New IO.StreamReader(programConstants.configXMLFile)
                    Dim xmlSerializerObject As New XmlSerializer(AppSettingsObject.GetType)
                    AppSettingsObject = xmlSerializerObject.Deserialize(streamReader)
                End Using
            End SyncLock
        End Sub

        Public Sub SaveSettingsToXMLFile()
            SyncLock programFunctions.LockObject
                Using streamWriter As New IO.StreamWriter(programConstants.configXMLFile)
                    Dim xmlSerializerObject As New XmlSerializer(AppSettingsObject.GetType)
                    xmlSerializerObject.Serialize(streamWriter, AppSettingsObject)
                End Using
            End SyncLock
        End Sub
    End Module
End Namespace