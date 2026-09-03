' ============================================================================
'  Modules\ProjectModel.vb
'
'  What a .Busmaster project file contains: enough to put the workspace back the
'  way it was, and nothing else for now.
'
'  Devices are written in the order they sit on screen, so reading them back in
'  order restores the layout.
' ============================================================================

Public Class ProjectData

    Public Property SchemaVersion As Integer = 1
    Public Property ProjectName As String = "Untitled"
    Public Property CreatedUtc As String = String.Empty
    Public Property ModifiedUtc As String = String.Empty

    Public Property Devices As New List(Of ProjectDeviceData)

End Class


''' <summary>One device panel: what it is called, and which part it holds.</summary>
Public Class ProjectDeviceData

    Public Property FunctionName As String = String.Empty

    ''' <summary>Device name without path or extension - "PCA9555".</summary>
    Public Property DeviceFile As String = String.Empty

End Class
