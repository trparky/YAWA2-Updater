Option Strict Off
Namespace internetFunctions
    Module Internet_Functions
        Public Function createNewHTTPHelperObject() As httpHelper
            Dim httpHelper As New httpHelper With {
                .setUserAgent = createHTTPUserAgentHeaderString(),
                .useHTTPCompression = True,
                .setProxyMode = True
            }
            httpHelper.addHTTPHeader("OPERATING_SYSTEM", getFullOSVersionString())

            httpHelper.setURLPreProcessor = Function(ByVal strURLInput As String) As String
                                                Try
                                                    If Not strURLInput.Trim.ToLower.StartsWith("http") Then
                                                        If boolUseSSL Then
                                                            Debug.WriteLine("setURLPreProcessor code -- https://" & strURLInput)
                                                            Return "https://" & strURLInput
                                                        Else
                                                            Debug.WriteLine("setURLPreProcessor code -- http://" & strURLInput)
                                                            Return "http://" & strURLInput
                                                        End If
                                                    Else
                                                        Debug.WriteLine("setURLPreProcessor code -- " & strURLInput)
                                                        Return strURLInput
                                                    End If
                                                Catch ex As Exception
                                                    Debug.WriteLine("setURLPreProcessor code -- " & strURLInput)
                                                    Return strURLInput
                                                End Try
                                            End Function

            Return httpHelper
        End Function

        ''' <summary>Creates a User Agent String for this program to be used in HTTP requests.</summary>
        ''' <returns>String type.</returns>
        Private Function createHTTPUserAgentHeaderString() As String
            Dim versionInfo As String() = Application.ProductVersion.Split(".")
            Dim versionString As String = String.Format("{0}.{1} Build {2}", versionInfo(0), versionInfo(1), versionInfo(2))
            Return String.Format("{0} version {1} on {2}", programName, versionString, getFullOSVersionString())
        End Function

        Private Function getFullOSVersionString() As String
            Try
                Dim computerInfo As New Devices.ComputerInfo()
                Dim osName As String = Text.RegularExpressions.Regex.Replace(computerInfo.OSFullName.Trim, "microsoft ", "", Text.RegularExpressions.RegexOptions.IgnoreCase)
                computerInfo = Nothing

                Dim dotNetVersionsInfo As String() = Environment.Version.ToString.Split(".")

                If Environment.Is64BitOperatingSystem Then
                    Return String.Format("{0} 64-bit (Microsoft .NET {1}.{2})", osName, dotNetVersionsInfo(0), dotNetVersionsInfo(1))
                Else
                    Return String.Format("{0} 32-bit (Microsoft .NET {1}.{2})", osName, dotNetVersionsInfo(0), dotNetVersionsInfo(1))
                End If
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