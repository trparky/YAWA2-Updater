Imports System.Xml.Serialization

Namespace internetFunctions
    Module Internet_Functions
        Public Function CreateNewHTTPHelperObject() As HttpHelper
            Dim httpHelper As New HttpHelper With {.SetUserAgent = CreateHTTPUserAgentHeaderString(), .UseHTTPCompression = True, .SetProxyMode = True}
            httpHelper.AddHTTPHeader("OPERATING_SYSTEM", GetFullOSVersionString())

            httpHelper.SetURLPreProcessor = Function(strURLInput As String) As String
                                                Try
                                                    If Not strURLInput.Trim.StartsWith("http", StringComparison.OrdinalIgnoreCase) Then
                                                        Return $"https://{strURLInput}"
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
            Dim versionString As String = $"{versionInfo(0)}.{versionInfo(1)} Build {versionInfo(2)}"
            Return $"{"YAWA2 (Yet Another WinApp2.ini) Updater"} version {versionString} on {GetFullOSVersionString()}"
        End Function

        Private Function GetFullOSVersionString() As String
            Try
                Dim osName As String = New Devices.ComputerInfo().OSFullName.Trim.CaseInsensitiveReplace("microsoft ", "", True)

                Dim dotNetVersionsInfo As String() = Environment.Version.ToString.Split(".")
                Return $"{osName} {If(Environment.Is64BitOperatingSystem, "64", "32")}-bit (Microsoft .NET {dotNetVersionsInfo(0)}.{dotNetVersionsInfo(1)})"
            Catch ex As Exception
                Try
                    Return $"Unknown Windows Operating System ({Environment.OSVersion.VersionString})"
                Catch ex2 As Exception
                    Return "Unknown Windows Operating System"
                End Try
            End Try
        End Function
    End Module
End Namespace