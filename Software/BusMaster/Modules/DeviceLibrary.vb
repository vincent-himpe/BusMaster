' ============================================================================
'  Modules\DeviceLibrary.vb
'
'  The device library on disk: the .\devices folder, the .DEV naming rule, and
'  reading and writing device files.
'
'  DeviceFolder sits beside the executable, the same place settings.json lives,
'  so a device library follows the installation rather than any one project.
'  Change that one property to move it.
' ============================================================================

Imports System.IO
Imports Microsoft.Win32

Public Module DeviceLibrary

    Public Const DeviceExtension As String = ".DEV"
    Public Const DeviceFolderName As String = "devices"

    Private Const DeviceFilter As String = "Device Files (*.DEV)|*.DEV|All Files (*.*)|*.*"

    ''' <summary>The .\devices folder. One property to relocate the whole library.</summary>
    Public ReadOnly Property DeviceFolder As String
        Get
            Return Path.Combine(AppContext.BaseDirectory, DeviceFolderName)
        End Get
    End Property

    ''' <summary>Creates the device folder if it is not there yet.</summary>
    Public Sub EnsureFolder()

        If Not Directory.Exists(DeviceFolder) Then Directory.CreateDirectory(DeviceFolder)

    End Sub

    ''' <summary>Where a device of this name belongs: .\devices\NAME.DEV</summary>
    Public Function PathFor(deviceName As String) As String

        Return Path.Combine(DeviceFolder, deviceName.Trim() & DeviceExtension)

    End Function

    ''' <summary>True when a device of this name is already in the library.</summary>
    Public Function DeviceExists(deviceName As String) As Boolean

        If String.IsNullOrWhiteSpace(deviceName) Then Return False

        Return File.Exists(PathFor(deviceName))

    End Function

    ''' <summary>
    ''' True when a name is letters and numbers only. That is stricter than Windows
    ''' needs for a file name, which is the point: it also rules out spaces, dots and
    ''' the reserved device names, so NAME.DEV is always a sane file.
    ''' </summary>
    Public Function IsUsableDeviceName(deviceName As String) As Boolean

        If String.IsNullOrWhiteSpace(deviceName) Then Return False

        For Each character As Char In deviceName.Trim()

            Dim allowed As Boolean =
                (character >= "A"c AndAlso character <= "Z"c) OrElse
                (character >= "a"c AndAlso character <= "z"c) OrElse
                (character >= "0"c AndAlso character <= "9"c)

            If Not allowed Then Return False

        Next

        Return True

    End Function

    ''' <summary>Reads a .DEV file. Throws; callers decide how to report it.</summary>
    Public Function LoadDeviceFile(filePath As String) As DeviceFileData

        Return JsonFile.ReadFrom(Of DeviceFileData)(filePath)

    End Function

    ''' <summary>Writes a .DEV file, creating .\devices first. Throws.</summary>
    Public Sub SaveDeviceFile(filePath As String, data As DeviceFileData)

        EnsureFolder()
        JsonFile.WriteTo(filePath, data)

    End Sub

    ''' <summary>
    ''' Asks the user for a device file, starting in the library folder. Returns an
    ''' empty string if the dialog was cancelled.
    ''' </summary>
    Public Function PickDeviceFile(owner As Window, title As String) As String

        EnsureFolder()

        Dim dialog As New OpenFileDialog With {
            .Title = title,
            .Filter = DeviceFilter,
            .DefaultExt = DeviceExtension,
            .CheckFileExists = True,
            .InitialDirectory = DeviceFolder
        }

        Dim picked As Boolean? = dialog.ShowDialog(owner)

        If Not picked.GetValueOrDefault() Then Return String.Empty

        Return dialog.FileName

    End Function

End Module
