Imports System.Runtime.CompilerServices
Imports System.Text.RegularExpressions

Module StringExtensions
    ' PHP like addSlashes and stripSlashes. Call using String.addSlashes() and String.stripSlashes().
    <Extension()>
    Public Function addSlashes(unsafeString As String) As String
        Return Regex.Replace(unsafeString, "([\000\010\011\012\015\032\042\047\134\140])", "\$1")
    End Function

    ''' <summary>This function uses an IndexOf call to do a case-insensitive search. This function operates a lot like Contains().</summary>
    ''' <param name="needle">The String containing what you want to search for.</param>
    ''' <return>Returns a Boolean value.</return>
    <Extension()>
    Public Function caseInsensitiveContains(haystack As String, needle As String) As Boolean
        Dim index As Integer = haystack.IndexOf(needle, StringComparison.OrdinalIgnoreCase)
        Return index <> -1
    End Function

    ''' <summary>This function uses a RegEx search to do a case-insensitive search but with more power and more responsibilities. This function operates a lot like Contains().</summary>
    ''' <param name="needle">The String containing what you want to search for.</param>
    ''' <param name="boolDoEscaping">This tells the function if it should add slashes where appropriate to the "needle" String.</param>
    ''' <return>Returns a Boolean value.</return>
    <Extension()>
    Public Function caseInsensitiveContains(haystack As String, needle As String, Optional boolDoEscaping As Boolean = False) As Boolean
        Try
            If boolDoEscaping Then needle = Regex.Escape(needle)
            Return Regex.IsMatch(haystack, needle, RegexOptions.IgnoreCase)
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>This function operates a lot like Replace() but is case-InSeNsItIvE.</summary>
    ''' <param name="source">The source String, aka the String where the data will be replaced in.</param>
    ''' <param name="replace">What you want to replace in the String.</param>
    ''' <param name="replaceWith">What you want to replace with in the String.</param>
    ''' <param name="boolEscape">This is an optional parameter, the default is True. This parameter gives you far more control over how the function works. With this parameter set to True the function automatically properly escapes the "replace" parameter for use in the RegEx replace function that operates inside this function. If this parameter is set to False it is up to you, the programmer, to properly escape the value of the "replace" parameter or this function will throw an exception.</param>
    ''' <return>Returns a String value.</return>
    <Extension()>
    Public Function caseInsensitiveReplace(source As String, replace As String, replaceWith As String, Optional boolEscape As Boolean = True) As String
        If boolEscape Then replace = Regex.Escape(replace)
        Return Regex.Replace(source, replace, replaceWith, RegexOptions.IgnoreCase)
    End Function
End Module