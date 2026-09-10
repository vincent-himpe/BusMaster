' ============================================================================
'  DeviceEditorWindow.xaml.vb
'
'  Event handlers ONLY, and the two constructors that record what the form was
'  opened for. All of the work lives in Modules\DeviceEditorCore.vb.
'
'  Open it with no argument for a blank form, or with a file to load it:
'
'      New DeviceEditorWindow()
'      New DeviceEditorWindow(path, DeviceEditorMode.mode_Edit)
'      New DeviceEditorWindow(path, DeviceEditorMode.mode_Clone)
' ============================================================================

Imports System.ComponentModel

Class DeviceEditorWindow

    ''' <summary>What this instance was opened to do.</summary>
    Public Property EditorMode As DeviceEditorMode = DeviceEditorMode.mode_New

    ''' <summary>The file to load, or an empty string for a blank form.</summary>
    Public Property DeviceFilePath As String = String.Empty

    ''' <summary>Blank form, ready for a new device.</summary>
    Public Sub New()
        Me.New(String.Empty, DeviceEditorMode.mode_New)
    End Sub

    ''' <summary>Form loaded from a device file.</summary>
    Public Sub New(devicefile As String, mode As DeviceEditorMode)

        InitializeComponent()

        DeviceFilePath = If(devicefile, String.Empty)
        EditorMode = mode

    End Sub

    ' -- Window ---------------------------------------------------------------

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        DeviceEditorCore.Attach(Me)
    End Sub

    Private Sub Window_Closing(sender As Object, e As CancelEventArgs)
        DeviceEditorCore.HandleWindowClosing(e)
    End Sub

    ' -- Header fields --------------------------------------------------------

    Private Sub txt_DeviceName_TextChanged(sender As Object, e As TextChangedEventArgs)
        DeviceEditorCore.DeviceNameEdited()
    End Sub

    Private Sub txt_DeviceDescription_TextChanged(sender As Object, e As TextChangedEventArgs)
        DeviceEditorCore.DescriptionEdited()
    End Sub

    Private Sub txt_BaseAddress_TextChanged(sender As Object, e As TextChangedEventArgs)
        DeviceEditorCore.BaseAddressEdited()
    End Sub

    Private Sub txt_AddressBits_TextChanged(sender As Object, e As TextChangedEventArgs)
        DeviceEditorCore.AddressBitsEdited()
    End Sub

    ' -- Register table tool bar ----------------------------------------------

    Private Sub tlbr_Add8_Click(sender As Object, e As RoutedEventArgs)
        DeviceEditorCore.AddByteRegister()
    End Sub

    Private Sub tlbr_Add16_Click(sender As Object, e As RoutedEventArgs)
        DeviceEditorCore.AddWordRegister()
    End Sub

    Private Sub tlbr_DeleteRow_Click(sender As Object, e As RoutedEventArgs)
        DeviceEditorCore.DeleteCurrentRow()
    End Sub

    Private Sub tlbr_SortRows_Click(sender As Object, e As RoutedEventArgs)
        DeviceEditorCore.SortRowsByAddress()
    End Sub

    ' -- Buttons --------------------------------------------------------------

    Private Sub btn_Save_Click(sender As Object, e As RoutedEventArgs)
        DeviceEditorCore.SaveDevice()
    End Sub

    Private Sub btn_Cancel_Click(sender As Object, e As RoutedEventArgs)
        DeviceEditorCore.CancelEditor()
    End Sub

End Class
