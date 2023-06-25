Imports System.Runtime.CompilerServices
Imports System.Text.RegularExpressions

Module StringExtensions
    Public Const DoubleCRLF As String = vbCrLf & vbCrLf

    ' PHP like addSlashes and stripSlashes. Call using String.addSlashes() and String.stripSlashes().
    <Extension()>
    Public Function AddSlashes(unsafeString As String) As String
        Return Regex.Replace(unsafeString, "([\000\010\011\012\015\032\042\047\134\140])", "\$1")
    End Function

    ''' <summary>This function uses an IndexOf call to do a case-insensitive search. This function operates a lot like Contains().</summary>
    ''' <param name="needle">The String containing what you want to search for.</param>
    ''' <return>Returns a Boolean value.</return>
    <Extension()>
    Public Function CaseInsensitiveContains(haystack As String, needle As String) As Boolean
        Dim index As Integer = haystack.IndexOf(needle, StringComparison.OrdinalIgnoreCase)
        Return index <> -1
    End Function

    ''' <summary>This function uses a RegEx search to do a case-insensitive search but with more power and more responsibilities. This function operates a lot like Contains().</summary>
    ''' <param name="needle">The String containing what you want to search for.</param>
    ''' <param name="boolDoEscaping">This tells the function if it should add slashes where appropriate to the "needle" String.</param>
    ''' <return>Returns a Boolean value.</return>
    <Extension()>
    Public Function CaseInsensitiveContains(haystack As String, needle As String, Optional boolDoEscaping As Boolean = False) As Boolean
        Try
            If boolDoEscaping Then needle = Regex.Escape(needle)
            Return Regex.IsMatch(haystack, needle, RegexOptions.IgnoreCase)
        Catch ex As Exception
            Return False
        End Try
    End Function

    '''<summary>Works similar to the original String Replacement function but with a potential case-insensitive match capability.</summary>
    ''' <param name="source">The source String.</param>
    ''' <param name="strReplace">The String to be replaced.</param>
    ''' <param name="strReplaceWith">The String that will replace all occurrences of <paramref name="strReplaceWith"/>. 
    ''' If value Is equal to <c>null</c>, than all occurrences of <paramref name="strReplace"/> will be removed from the <paramref name="source"/>.</param>
    ''' <param name="comparisonType">One of the enumeration values that specifies the rules for the search.</param>
    ''' <returns>A string that Is equivalent to the current string except that all instances of <paramref name="strReplace"/> are replaced with <paramref name="strReplaceWith"/>. 
    ''' If <paramref name="strReplace"/> Is Not found in the current instance, the method returns the current instance unchanged.</returns>
    <Extension()>
    Public Function Replace(source As String, strReplace As String, strReplaceWith As String, comparisonType As StringComparison) As String
        If String.IsNullOrWhiteSpace(source) Then Throw New ArgumentNullException(NameOf(source))
        If source.Length = 0 Then Return source
        If String.IsNullOrWhiteSpace(strReplace) Then Throw New ArgumentNullException(NameOf(strReplace))
        If strReplace.Length = 0 Then Throw New ArgumentException("String cannot be of zero length.")

        Dim resultStringBuilder As New Text.StringBuilder(source.Length)
        Dim isReplacementNullOrEmpty As Boolean = String.IsNullOrEmpty(strReplaceWith)

        Const valueNotFound As Integer = -1
        Dim foundAt As Integer
        Dim startSearchFromIndex As Integer = 0

        While InlineAssignHelper(foundAt, source.IndexOf(strReplace, startSearchFromIndex, comparisonType)) <> valueNotFound
            Dim charsUntilReplacment As Integer = foundAt - startSearchFromIndex
            Dim isNothingToAppend As Boolean = charsUntilReplacment = 0

            If Not isNothingToAppend Then resultStringBuilder.Append(source, startSearchFromIndex, charsUntilReplacment)
            If Not isReplacementNullOrEmpty Then resultStringBuilder.Append(strReplaceWith)

            startSearchFromIndex = foundAt + strReplace.Length
            If startSearchFromIndex = source.Length Then Return resultStringBuilder.ToString()
        End While

        Dim charsUntilStringEnd As Integer = source.Length - startSearchFromIndex
        resultStringBuilder.Append(source, startSearchFromIndex, charsUntilStringEnd)

        Return resultStringBuilder.ToString()
    End Function

    Private Function InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
        target = value
        Return value
    End Function
End Module