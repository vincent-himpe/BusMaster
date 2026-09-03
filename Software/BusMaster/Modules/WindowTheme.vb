' ============================================================================
'  Modules\WindowTheme.vb
'
'  The window frame - title bar, caption buttons - is drawn by Windows, not by
'  WPF, so a dark application still gets a light title bar unless we ask DWM for
'  the dark one. This is the same call Visual Studio makes.
'
'  Best effort only: on an OS that does not support the attribute the call fails
'  and the window simply keeps the default frame.
' ============================================================================

Imports System.Runtime.InteropServices
Imports System.Windows.Interop

Public Module WindowTheme

    ' Windows 10 build 18985 and later.
    Private Const DWMWA_USE_IMMERSIVE_DARK_MODE As Integer = 20

    ' Windows 10 builds 17763 - 18984 used this value instead.
    Private Const DWMWA_USE_IMMERSIVE_DARK_MODE_LEGACY As Integer = 19

    <DllImport("dwmapi.dll", PreserveSig:=True)>
    Private Function DwmSetWindowAttribute(hwnd As IntPtr,
                                           attribute As Integer,
                                           ByRef value As Integer,
                                           valueSize As Integer) As Integer
    End Function

    ''' <summary>
    ''' Switches a window's title bar to the dark variant. Must be called once the
    ''' window has a handle, so from Loaded or later.
    ''' </summary>
    Public Sub ApplyDarkTitleBar(window As Window)

        If window Is Nothing Then Exit Sub

        Try
            Dim handle As IntPtr = New WindowInteropHelper(window).Handle
            If handle = IntPtr.Zero Then Exit Sub

            Dim enabled As Integer = 1

            If DwmSetWindowAttribute(handle, DWMWA_USE_IMMERSIVE_DARK_MODE, enabled, 4) <> 0 Then
                DwmSetWindowAttribute(handle, DWMWA_USE_IMMERSIVE_DARK_MODE_LEGACY, enabled, 4)
            End If

        Catch
            ' Cosmetic only - never let this stop the program starting.
        End Try

    End Sub

End Module
