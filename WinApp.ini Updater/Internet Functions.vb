﻿Namespace internetFunctions
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
                                                        Return If(boolUseSSL, "https://", "http://") & strURLInput
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