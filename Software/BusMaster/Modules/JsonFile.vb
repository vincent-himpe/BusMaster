' ============================================================================
'  Modules\JsonFile.vb
'
'  One place that decides how this program reads and writes JSON, so settings
'  files and project files behave identically:
'
'    - written indented, so a human can read and diff them
'    - read tolerantly, so comments and trailing commas survive hand editing
'
'  Uses System.Text.Json from the .NET framework itself - no NuGet packages.
' ============================================================================

Imports System.IO
Imports System.Text.Json

Public Module JsonFile

    Friend ReadOnly ReadOptions As New JsonSerializerOptions With {
        .ReadCommentHandling = JsonCommentHandling.Skip,
        .AllowTrailingCommas = True,
        .PropertyNameCaseInsensitive = True
    }

    Friend ReadOnly WriteOptions As New JsonSerializerOptions With {
        .WriteIndented = True
    }

    ''' <summary>
    ''' Reads a JSON file into a new object. Throws if the file is missing or
    ''' malformed; a file containing literal "null" yields a default object.
    ''' </summary>
    Public Function ReadFrom(Of T As {Class, New})(filePath As String) As T

        Dim json As String = File.ReadAllText(filePath)

        If String.IsNullOrWhiteSpace(json) Then Return New T

        Dim value As T = JsonSerializer.Deserialize(Of T)(json, ReadOptions)
        Return If(value, New T)

    End Function

    ''' <summary>
    ''' Writes an object to a JSON file, creating the directory if it is missing.
    ''' Throws on failure - callers decide how to report it.
    ''' </summary>
    Public Sub WriteTo(Of T)(filePath As String, value As T)

        Dim folder As String = Path.GetDirectoryName(filePath)

        If Not String.IsNullOrEmpty(folder) AndAlso Not Directory.Exists(folder) Then
            Directory.CreateDirectory(folder)
        End If

        File.WriteAllText(filePath, JsonSerializer.Serialize(value, WriteOptions))

    End Sub

End Module
