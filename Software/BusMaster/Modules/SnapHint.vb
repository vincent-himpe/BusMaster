' ============================================================================
'  Modules\SnapHint.vb
'
'  The visual cue that a tool window being dragged is near enough to an edge of
'  the main window to stick to it.
'
'  Without it, snapping is guesswork. The reach is twenty pixels, the decision is
'  taken at the moment the mouse comes up, and until then nothing on screen says
'  whether you are inside it - so a window that does not snap looks exactly like
'  a window that was never going to. The overlay answers the only question the
'  user actually has: let go now, and does it take?
'
'  While an edge is live the main window is greyed and outlined in light blue,
'  and the edge that will take the window is drawn as a thicker bar with a glow
'  fading inwards from it. The overlay appears only while an edge is live, so its
'  presence alone is the answer and the bar says which side.
'
'  Drawn as an adorner over the main window's content, the same way ModalShade
'  works and for the same reason: the main window knows nothing about any of this
'  and does not have to. The differences are that this wash is lighter - a modal
'  shade means stop and answer me, this one means keep going, you are nearly
'  there - and that it is only ever the one window, never the whole application.
'
'  The glow is a gradient drawn into the adorner rather than a BlurEffect on it,
'  because an effect applies to the whole visual: it would take the crisp outline
'  and the bar with it and leave the lot looking out of focus.
' ============================================================================

Imports System.Windows.Documents
Imports System.Windows.Media

Public Module SnapHint

    Private Const ShadeBrushKey As String = "Brush_Snap_Shade"
    Private Const EdgeBrushKey As String = "Brush_Snap_Edge"


    Private Sheet As SnapHintAdorner


    ''' <summary>
    ''' Shows the cue on the main window with one edge live, or moves the live edge
    ''' if it is already up. Called on every mouse move of a drag, so being asked
    ''' for what is already on screen is the normal case, not the exception.
    ''' </summary>
    Public Sub Show(main As Window, edge As SnapEdge)

        If main Is Nothing OrElse Not main.IsVisible Then
            Hide()
            Exit Sub
        End If

        If edge = SnapEdge.edge_None Then
            Hide()
            Exit Sub
        End If

        If Sheet IsNot Nothing AndAlso Sheet.Window IsNot main Then Hide()

        If Sheet Is Nothing Then Sheet = Raise(main)
        If Sheet Is Nothing Then Exit Sub

        Sheet.Live = edge

    End Sub

    ''' <summary>Takes the cue down. Safe to call when there is none up.</summary>
    Public Sub Hide()

        If Sheet Is Nothing Then Exit Sub

        Dim layer As AdornerLayer = AdornerLayer.GetAdornerLayer(Sheet.AdornedElement)
        If layer IsNot Nothing Then layer.Remove(Sheet)

        Sheet = Nothing

    End Sub

    ''' <summary>
    ''' Puts a fresh sheet on the window, or Nothing if there is nothing to draw on
    ''' or the theme is missing its colours.
    ''' </summary>
    Private Function Raise(main As Window) As SnapHintAdorner

        Dim root As UIElement = TryCast(main.Content, UIElement)
        If root Is Nothing Then Return Nothing

        Dim layer As AdornerLayer = AdornerLayer.GetAdornerLayer(root)
        If layer Is Nothing Then Return Nothing

        Dim shade As Brush = TryCast(main.TryFindResource(ShadeBrushKey), Brush)
        Dim accent As SolidColorBrush = TryCast(main.TryFindResource(EdgeBrushKey), SolidColorBrush)

        If shade Is Nothing OrElse accent Is Nothing Then Return Nothing

        Dim made As New SnapHintAdorner(root, main, shade, accent.Color)
        layer.Add(made)

        Return made

    End Function

End Module


''' <summary>
''' The wash, the outline and the one live edge. Not hit testable: the main window
''' is not disabled while this is up - the user is dragging a different window and
''' may let go anywhere - so nothing here may swallow a click.
''' </summary>
Friend NotInheritable Class SnapHintAdorner
    Inherits Adorner

    ''' <summary>The outline right round the window, saying which window will take it.</summary>
    Private Const OutlineThickness As Double = 2

    ''' <summary>The live edge, drawn heavier than the outline to pick it out.</summary>
    Private Const EdgeThickness As Double = 6

    ''' <summary>How far the glow reaches in from the live edge before it is gone.</summary>
    Private Const GlowDepth As Double = 22

    ''' <summary>How solid the glow is where it meets the edge.</summary>
    Private Const GlowStrength As Byte = &H70


    Private ReadOnly Shade As Brush
    Private ReadOnly Accent As Brush
    Private ReadOnly Outline As Pen
    Private ReadOnly GlowNear As Color
    Private ReadOnly GlowFar As Color

    ''' <summary>The window this is drawn on, so a later call can tell it is the same one.</summary>
    Public ReadOnly Window As Window

    Private Edge As SnapEdge = SnapEdge.edge_None


    Public Sub New(over As UIElement, owner As Window, wash As Brush, ink As Color)

        MyBase.New(over)

        Window = owner
        Shade = wash

        Accent = New SolidColorBrush(ink)
        Accent.Freeze()

        Outline = New Pen(Accent, OutlineThickness)
        Outline.Freeze()

        GlowNear = Color.FromArgb(GlowStrength, ink.R, ink.G, ink.B)
        GlowFar = Color.FromArgb(0, ink.R, ink.G, ink.B)

        IsHitTestVisible = False
        SnapsToDevicePixels = True

    End Sub


    ''' <summary>
    ''' Which edge is about to take the window. Repaints only on a change, because
    ''' this is written on every mouse move of a drag.
    ''' </summary>
    Public Property Live As SnapEdge
        Get
            Return Edge
        End Get
        Set(value As SnapEdge)
            If value = Edge Then Exit Property
            Edge = value
            InvalidateVisual()
        End Set
    End Property


    Protected Overrides Sub OnRender(drawing As DrawingContext)

        Dim area As New Rect(AdornedElement.RenderSize)
        If area.Width <= 0 OrElse area.Height <= 0 Then Exit Sub

        drawing.DrawRectangle(Shade, Nothing, area)

        ' A pen straddles the line it is given, so half of a border drawn on the
        ' outline of the window falls outside it and is clipped away.
        drawing.DrawRectangle(Nothing, Outline,
                              Rect.Inflate(area, -OutlineThickness / 2, -OutlineThickness / 2))

        PaintEdge(drawing, area)

    End Sub

    ''' <summary>The live edge: a bar on the edge itself and a glow fading inwards.</summary>
    Private Sub PaintEdge(drawing As DrawingContext, area As Rect)

        If Edge = SnapEdge.edge_None Then Exit Sub

        Dim bar As Rect
        Dim glow As Rect
        Dim fromEnd As Point
        Dim toEnd As Point

        ' A window narrow enough for the glow to reach its middle gets a shorter
        ' glow rather than two that meet.
        Dim depth As Double = Math.Max(Math.Min(area.Width, area.Height) / 2 - EdgeThickness, 0)
        depth = Math.Min(GlowDepth, depth)

        Select Case Edge

            Case SnapEdge.edge_Bottom
                bar = New Rect(area.Left, area.Bottom - EdgeThickness, area.Width, EdgeThickness)
                glow = New Rect(area.Left, bar.Top - depth, area.Width, depth)
                fromEnd = New Point(0, 1)
                toEnd = New Point(0, 0)

            Case SnapEdge.edge_Top
                bar = New Rect(area.Left, area.Top, area.Width, EdgeThickness)
                glow = New Rect(area.Left, bar.Bottom, area.Width, depth)
                fromEnd = New Point(0, 0)
                toEnd = New Point(0, 1)

            Case SnapEdge.edge_Right
                bar = New Rect(area.Right - EdgeThickness, area.Top, EdgeThickness, area.Height)
                glow = New Rect(bar.Left - depth, area.Top, depth, area.Height)
                fromEnd = New Point(1, 0)
                toEnd = New Point(0, 0)

            Case SnapEdge.edge_Left
                bar = New Rect(area.Left, area.Top, EdgeThickness, area.Height)
                glow = New Rect(bar.Right, area.Top, depth, area.Height)
                fromEnd = New Point(0, 0)
                toEnd = New Point(1, 0)

        End Select

        If depth > 0 Then

            Dim fade As New LinearGradientBrush(GlowNear, GlowFar, fromEnd, toEnd)
            fade.Freeze()

            drawing.DrawRectangle(fade, Nothing, glow)

        End If

        drawing.DrawRectangle(Accent, Nothing, bar)

    End Sub

End Class
