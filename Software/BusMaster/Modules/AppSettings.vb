' ============================================================================
'  Modules\AppSettings.vb
'
'  Program settings - window geometry, options and the most-recently-used list.
'
'  Persisted as settings.json beside the executable. JSON was chosen over TOML
'  because System.Text.Json ships with .NET, so the whole file is one call in and
'  one call out and the project needs no NuGet packages at all.
'
'  Everything in SettingsData is deliberately a String, a number, a Boolean or a
'  list of Strings, so the file stays readable and hand-editable.
' ============================================================================

Imports System.IO

''' <summary>
''' The shape of settings.json. Property names are the JSON key names.
''' </summary>
Public Class SettingsData
    Public Property SchemaVersion As Integer = 1
    Public Property WindowWidth As Double = 1000
    Public Property WindowHeight As Double = 650
    Public Property WindowMaximized As Boolean = False
    Public Property ConfirmOnExit As Boolean = False
    Public Property MaxRecentProjects As Integer = 5
    Public Property LastProjectFolder As String = String.Empty
    Public Property RecentProjects As New List(Of String)
End Class

Public Module AppSettings

    ''' <summary>
    ''' Program version. Shown in the title bar and the help box. This is the one
    ''' place it is written down.
    ''' </summary>
    Public Const ApplicationVersion As String = "0.1"

    Public Const SettingsFileName As String = "settings.json"

    ''' <summary>
    ''' The two squares of the workspace checkerboard. Light is 75% black, dark is
    ''' 90% black; the tile is drawn 3 pixels to a square.
    ''' </summary>
    Public Const Checker_Light As String = "#FF404040"
    Public Const Checker_Dark As String = "#FF1A1A1A"

    ''' <summary>Side of one checkerboard square, in pixels.</summary>
    Public Const Checker_Size As Double = 3

    ''' <summary>
    ''' Upper bound for MaxRecentProjects. The File menu declares one fixed slot per
    ''' entry (mnu_File_Recent_1 .. mnu_File_Recent_5), so raising this means adding
    ''' matching menu items to MainWindow.xaml.
    ''' </summary>
    Public Const RecentProjectSlots As Integer = 5

    ''' <summary>The live settings. Never Nothing.</summary>
    Public Property Current As New SettingsData

    ''' <summary>Message from the last failed Load or Save, or an empty string.</summary>
    Public Property LastError As String = String.Empty

    ''' <summary>
    ''' Full path of settings.json. Change this one property to relocate the file
    ''' (to %APPDATA%, for example).
    ''' </summary>
    Public ReadOnly Property SettingsFilePath As String
        Get
            Return Path.Combine(AppContext.BaseDirectory, SettingsFileName)
        End Get
    End Property

    ''' <summary>
    ''' Loads settings.json. A missing file is not an error - defaults are used and
    ''' the file is created. A corrupt file falls back to defaults and is left on
    ''' disk untouched so it can be inspected.
    ''' </summary>
    Public Function Load() As Boolean

        LastError = String.Empty

        Try
            If Not File.Exists(SettingsFilePath) Then
                Current = New SettingsData
                Return Save()
            End If

            Current = JsonFile.ReadFrom(Of SettingsData)(SettingsFilePath)
            Normalise()
            Return True

        Catch ex As Exception
            Current = New SettingsData
            LastError = ex.Message
            Return False
        End Try

    End Function

    ''' <summary>Writes the live settings back to settings.json.</summary>
    Public Function Save() As Boolean

        LastError = String.Empty

        Try
            Normalise()
            JsonFile.WriteTo(SettingsFilePath, Current)
            Return True

        Catch ex As Exception
            LastError = ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' Puts a project at the top of the MRU list, removing any earlier entry for
    ''' the same file and trimming the list to MaxRecentProjects.
    ''' </summary>
    Public Sub AddRecentProject(projectPath As String)

        If String.IsNullOrWhiteSpace(projectPath) Then Exit Sub

        Dim fullPath As String

        Try
            fullPath = Path.GetFullPath(projectPath)
        Catch
            fullPath = projectPath
        End Try

        Current.RecentProjects.RemoveAll(
            Function(entry As String) String.Equals(entry, fullPath, StringComparison.OrdinalIgnoreCase))

        Current.RecentProjects.Insert(0, fullPath)

        Save()

    End Sub

    ''' <summary>Drops a project from the MRU list - used when a file has gone missing.</summary>
    Public Sub RemoveRecentProject(projectPath As String)

        If String.IsNullOrWhiteSpace(projectPath) Then Exit Sub

        Current.RecentProjects.RemoveAll(
            Function(entry As String) String.Equals(entry, projectPath, StringComparison.OrdinalIgnoreCase))

        Save()

    End Sub

    ''' <summary>
    ''' Repairs anything a hand-edited or out-of-date file might contain, so the
    ''' rest of the program can trust these values.
    ''' </summary>
    Private Sub Normalise()

        If Current Is Nothing Then Current = New SettingsData

        If Current.MaxRecentProjects < 1 Then Current.MaxRecentProjects = 1
        If Current.MaxRecentProjects > RecentProjectSlots Then Current.MaxRecentProjects = RecentProjectSlots

        ' Must not fall below the window's own MinWidth / MinHeight.
        If Current.WindowWidth < 600 Then Current.WindowWidth = 600
        If Current.WindowHeight < 600 Then Current.WindowHeight = 600

        If Current.LastProjectFolder Is Nothing Then Current.LastProjectFolder = String.Empty

        If Current.RecentProjects Is Nothing Then
            Current.RecentProjects = New List(Of String)
        Else
            Current.RecentProjects.RemoveAll(Function(entry As String) String.IsNullOrWhiteSpace(entry))

            If Current.RecentProjects.Count > Current.MaxRecentProjects Then
                Current.RecentProjects.RemoveRange(
                    Current.MaxRecentProjects,
                    Current.RecentProjects.Count - Current.MaxRecentProjects)
            End If
        End If

    End Sub

End Module
