Imports System.Xml.Serialization

Namespace internetFunctions
    Module Internet_Functions
        Public Function CreateNewHTTPHelperObject() As HttpHelper
            Dim httpHelper As New HttpHelper With {.SetUserAgent = CreateHTTPUserAgentHeaderString(), .UseHTTPCompression = True, .SetProxyMode = True}
            httpHelper.AddHTTPHeader("OPERATING_SYSTEM", GetFullOSVersionString())

            Dim AppSettings As New AppSettings

            Using streamReader As New IO.StreamReader(programConstants.configXMLFile)
                Dim xmlSerializerObject As New XmlSerializer(AppSettings.GetType)
                AppSettings = xmlSerializerObject.Deserialize(streamReader)
            End Using

            httpHelper.SetURLPreProcessor = Function(ByVal strURLInput As String) As String
                                                Try
                                                    If Not strURLInput.Trim.ToLower.StartsWith("http") Then
                                                        Return If(AppSettings.boolUseSSL, "https://", "http://") & strURLInput
                                                    Else
                                                        Return strURLInput
                                                    End If
                                                Catch ex As Exception
                                                    Return strURLInput
                                                End Try
                                            End Function

            Return httpHelper
        End Function

        ''' <summary>Creates a User Agent String for this program to be used in HTTP requests.</summary>
        ''' <returns>String type.</returns>
        Private Function CreateHTTPUserAgentHeaderString() As String
            Dim versionInfo As String() = Application.ProductVersion.Split(".")
            Dim versionString As String = String.Format("{0}.{1} Build {2}", versionInfo(0), versionInfo(1), versionInfo(2))
            Return String.Format("{0} version {1} on {2}", "YAWA2 (Yet Another WinApp2.ini) Updater", versionString, GetFullOSVersionString())
        End Function

        Private Function GetFullOSVersionString() As String
            Try
                Dim osName As String = New Devices.ComputerInfo().OSFullName.Trim.CaseInsensitiveReplace("microsoft ", "", True)

                Dim dotNetVersionsInfo As String() = Environment.Version.ToString.Split(".")
                Return String.Format("{0} {3}-bit (Microsoft .NET {1}.{2})", osName, dotNetVersionsInfo(0), dotNetVersionsInfo(1), If(Environment.Is64BitOperatingSystem, "64", "32"))
            Catch ex As Exception
                Try
                    Return "Unknown Windows Operating System (" & Environment.OSVersion.VersionString & ")"
                Catch ex2 As Exception
                    Return "Unknown Windows Operating System"
                End Try
            End Try
        End Function
    End Module
End Namespace