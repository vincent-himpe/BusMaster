' ============================================================================
'  Modules\WorkspaceModel.vb
'
'  What a "<project>.Workspace" file holds: a set of named views, each one a
'  record of which registers had their padlock shut when it was saved.
'
'  Kept in its own file, beside the project rather than inside it, so views can
'  be added, thrown away or handed to someone else without touching the layout
'  the project describes.
'
'  Panels are found by the identity they were given when they were added, not by
'  their position or their name, so a view still lands correctly after the user
'  has reordered, renamed or reloaded them. A panel the view has never heard of
'  is left alone; a panel the view names but the project no longer has is
'  skipped. Neither is an error.
' ============================================================================

''' <summary>The whole file - every view saved against one project.</summary>
Public Class WorkspaceFile

    Public Property SchemaVersion As Integer = 1
    Public Property Views As New List(Of WorkspaceView)

End Class


''' <summary>One named view, as it appears in the toolbar's Workspace list.</summary>
Public Class WorkspaceView

    Public Property Name As String = String.Empty
    Public Property Panels As New List(Of WorkspacePanelState)

End Class


''' <summary>What one device panel looked like when the view was saved.</summary>
Public Class WorkspacePanelState

    ''' <summary>The panel's identity, from the project file.</summary>
    Public Property PanelId As String = String.Empty

    ''' <summary>
    ''' Addresses of the registers whose padlock was shut. Addresses rather than
    ''' positions, so the view survives a device file gaining or losing registers.
    ''' </summary>
    Public Property LockedRegisters As New List(Of Integer)

End Class
