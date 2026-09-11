' ============================================================================
'  Modules\ModalShade.vb
'
'  Lays a translucent sheet over every window that is waiting while a modal
'  dialog is up.
'
'  ShowDialog already disables the rest of the program - every other window of
'  the application stops answering the mouse and the keyboard. What it does not
'  do is make them look it: they stay in full colour and full contrast, so
'  clicking one and getting nothing reads as the program ignoring you rather
'  than as the program waiting for you. The sheet is the missing signal.
'
'  Drawn as an adorner rather than as something added to the window's own layout,
'  because it must cover whatever that window happens to hold without any of
'  those windows knowing about it. Every WPF window's default template carries an
'  AdornerDecorator, so there is always a layer to draw on.
'
'  Cover and Uncover nest. A dialog can put up a dialog of its own - the Device
'  Editor asking whether to overwrite a file - and the second one covers the
'  editor without laying a second sheet over the windows already covered.
' ============================================================================

Imports System.Windows.Documents
Imports System.Windows.Media

Public Module ModalShade

    Private Const ShadeBrushKey As String = "Brush_Modal_Shade"

    ''' <summary>The sheets each Cover put up, so Uncover takes down its own.</summary>
    Private ReadOnly Layers As New Stack(Of List(Of ShadeAdorner))

    ''' <summary>Windows already covered, so a nested dialog does not darken twice.</summary>
    Private ReadOnly Covered As New HashSet(Of Window)


    ''' <summary>
    ''' Covers every visible window of the application except the one about to be
    ''' shown. Call it immediately before ShowDialog.
    ''' </summary>
    Public Sub Cover(dialog As Window)

        Dim sheets As New List(Of ShadeAdorner)

        If Application.Current IsNot Nothing Then

            For Each window As Window In Application.Current.Windows

                If window Is Nothing OrElse window Is dialog Then Continue For
                If Not window.IsVisible Then Continue For
                If Covered.Contains(window) Then Continue For

                Dim sheet As ShadeAdorner = CoverOne(window)
                If sheet Is Nothing Then Continue For

                sheets.Add(sheet)
                Covered.Add(window)

            Next

        End If

        Layers.Push(sheets)

    End Sub

    ''' <summary>Takes down the sheets the matching Cover put up.</summary>
    Public Sub Uncover()

        If Layers.Count = 0 Then Exit Sub

        For Each sheet As ShadeAdorner In Layers.Pop()

            Dim layer As AdornerLayer = AdornerLayer.GetAdornerLayer(sheet.AdornedElement)
            If layer IsNot Nothing Then layer.Remove(sheet)

            Covered.Remove(sheet.Window)

        Next

    End Sub

    ''' <summary>
    ''' One window's sheet, or Nothing if it has nothing to draw on - a window with
    ''' no content yet, which is what the tool windows are until they are first
    ''' shown.
    ''' </summary>
    Private Function CoverOne(window As Window) As ShadeAdorner

        Dim root As UIElement = TryCast(window.Content, UIElement)
        If root Is Nothing Then Return Nothing

        Dim layer As AdornerLayer = AdornerLayer.GetAdornerLayer(root)
        If layer Is Nothing Then Return Nothing

        Dim ink As Brush = TryCast(window.TryFindResource(ShadeBrushKey), Brush)
        If ink Is Nothing Then Return Nothing

        Dim sheet As New ShadeAdorner(root, window, ink)
        layer.Add(sheet)

        Return sheet

    End Function

End Module


''' <summary>
''' A flat wash over the whole of what it is put on. Not hit testable - the window
''' underneath is disabled anyway, and a sheet that swallowed clicks would stop
''' Windows flashing the dialog when one is aimed at it.
''' </summary>
Friend NotInheritable Class ShadeAdorner
    Inherits Adorner

    Private ReadOnly Ink As Brush

    ''' <summary>The window this sheet belongs to, so it can be forgotten later.</summary>
    Public ReadOnly Window As Window

    Public Sub New(over As UIElement, owner As Window, shade As Brush)

        MyBase.New(over)

        Window = owner
        Ink = shade
        IsHitTestVisible = False

    End Sub

    Protected Overrides Sub OnRender(drawing As DrawingContext)

        drawing.DrawRectangle(Ink, Nothing, New Rect(AdornedElement.RenderSize))

    End Sub

End Class
