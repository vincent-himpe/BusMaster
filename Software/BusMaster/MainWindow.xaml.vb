' ============================================================================
'  MainWindow.xaml.vb
'
'  Event handlers ONLY. Every handler below is a single call into a module.
'  No application logic lives in this file - see Modules\AppCore.vb.
'
'  This is what lets the menu, the toolbar and the keyboard all drive the same
'  code: three handlers, one subroutine.
' ============================================================================

Imports System.ComponentModel

Class MainWindow

    ' -- Window ---------------------------------------------------------------

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        AppCore.Attach(Me)
    End Sub

    Private Sub Window_Closing(sender As Object, e As CancelEventArgs)
        AppCore.HandleWindowClosing(e)
    End Sub

    ' -- File menu ------------------------------------------------------------

    Private Sub mnu_File_New_Project_Click(sender As Object, e As RoutedEventArgs)
        AppCore.NewProject()
    End Sub

    Private Sub mnu_File_Open_Project_Click(sender As Object, e As RoutedEventArgs)
        AppCore.OpenProject()
    End Sub

    Private Sub mnu_File_Save_Project_Click(sender As Object, e As RoutedEventArgs)
        AppCore.SaveProject()
    End Sub

    Private Sub mnu_File_Save_Project_As_Click(sender As Object, e As RoutedEventArgs)
        AppCore.SaveProjectAs()
    End Sub

    ' Shared by mnu_File_Recent_1 through mnu_File_Recent_5. AppCore reads the
    ' path back off the menu item that raised the event.
    Private Sub mnu_File_Recent_Click(sender As Object, e As RoutedEventArgs)
        AppCore.OpenRecentProject(sender)
    End Sub

    Private Sub mnu_File_Options_Click(sender As Object, e As RoutedEventArgs)
        AppCore.ShowOptions()
    End Sub

    Private Sub mnu_File_Exit_Click(sender As Object, e As RoutedEventArgs)
        AppCore.ExitApplication()
    End Sub

    ' -- Device menu ----------------------------------------------------------

    Private Sub mnu_Device_Add_Device_Click(sender As Object, e As RoutedEventArgs)
        DeviceActions.AddDevice()
    End Sub

    Private Sub mnu_Device_Remove_Device_Click(sender As Object, e As RoutedEventArgs)
        DeviceActions.RemoveDevice()
    End Sub

    ' -- Bus menu -------------------------------------------------------------

    Private Sub mnu_Bus_Scan_Bus_Click(sender As Object, e As RoutedEventArgs)
        BusActions.ScanBus()
    End Sub

    Private Sub mnu_Bus_Poll_Bus_Click(sender As Object, e As RoutedEventArgs)
        BusActions.PollBus()
    End Sub

    Private Sub mnu_Bus_Reset_Bus_Click(sender As Object, e As RoutedEventArgs)
        BusActions.ResetBus()
    End Sub

    ' -- Tools menu -----------------------------------------------------------

    Private Sub mnu_Tools_Create_Device_Click(sender As Object, e As RoutedEventArgs)
        DeviceActions.CreateDevice()
    End Sub

    Private Sub mnu_Tools_Edit_Device_Click(sender As Object, e As RoutedEventArgs)
        DeviceActions.EditDevice()
    End Sub

    Private Sub mnu_Tools_Clone_Device_Click(sender As Object, e As RoutedEventArgs)
        DeviceActions.CloneDevice()
    End Sub

    Private Sub mnu_Tools_Show_Log_Click(sender As Object, e As RoutedEventArgs)
        EventLogCore.ShowLog()
    End Sub

    Private Sub tlbr_ShowLog_Click(sender As Object, e As RoutedEventArgs)
        EventLogCore.ShowLog()
    End Sub

    Private Sub mnu_Tools_Show_Config_Files_Click(sender As Object, e As RoutedEventArgs)
        AppCore.ShowConfigFolder()
    End Sub

    Private Sub mnu_Tools_Show_Dev_Files_Click(sender As Object, e As RoutedEventArgs)
        AppCore.ShowDeviceFolder()
    End Sub

    Private Sub mnu_Tools_Show_Project_Files_Click(sender As Object, e As RoutedEventArgs)
        AppCore.ShowProjectFolder()
    End Sub

    ' -- Tool bar -------------------------------------------------------------

    Private Sub tlbr_Open_Click(sender As Object, e As RoutedEventArgs)
        AppCore.OpenProject()
    End Sub

    Private Sub tlbr_Save_Click(sender As Object, e As RoutedEventArgs)
        AppCore.SaveProject()
    End Sub

    Private Sub tlbr_Help_Click(sender As Object, e As RoutedEventArgs)
        AppCore.ShowHelp()
    End Sub

    Private Sub tlbr_Record_Checked(sender As Object, e As RoutedEventArgs)
        AppCore.RecordingChanged(True)
    End Sub

    Private Sub tlbr_Record_Unchecked(sender As Object, e As RoutedEventArgs)
        AppCore.RecordingChanged(False)
    End Sub

    ' -- View menu ------------------------------------------------------------

    Private Sub mnu_View_Decimal_Click(sender As Object, e As RoutedEventArgs)
        AppCore.ShowValuesAsDecimal()
    End Sub

    Private Sub mnu_View_Hexadecimal_Click(sender As Object, e As RoutedEventArgs)
        AppCore.ShowValuesAsHexadecimal()
    End Sub

    ' -- Workspace ------------------------------------------------------------
    '    These two catch the bubbling events from every device on screen.

    Private Sub stk_Devices_IllegalEntry(sender As Object, e As RoutedEventArgs)
        AppCore.ReportIllegalData()
    End Sub

    Private Sub stk_Devices_LoadFailed(sender As Object, e As RoutedEventArgs)
        AppCore.ReportDeviceFileMissing()
    End Sub

    Private Sub stk_Devices_CloneRequested(sender As Object, e As RoutedEventArgs)
        AppCore.CloneDevicePanel(e.OriginalSource)
    End Sub

    Private Sub stk_Devices_DeleteRequested(sender As Object, e As RoutedEventArgs)
        AppCore.RemoveDevicePanel(e.OriginalSource)
    End Sub

    Private Sub tlbr_ValueType_Click(sender As Object, e As RoutedEventArgs)
        AppCore.ToggleValueDisplay()
    End Sub

    Private Sub tlbr_CollapseAll_Checked(sender As Object, e As RoutedEventArgs)
        AppCore.SetAllDevicesExpanded(True)
    End Sub

    Private Sub tlbr_CollapseAll_Unchecked(sender As Object, e As RoutedEventArgs)
        AppCore.SetAllDevicesExpanded(False)
    End Sub

    Private Sub tlbr_Operation_List_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        AppCore.OperationModeChosen(tlbr_Operation_List.SelectedIndex)
    End Sub

    Private Sub tlbr_Write_All_Click(sender As Object, e As RoutedEventArgs)
        BusActions.WriteAll()
    End Sub

    Private Sub tlbr_Workspace_Save_Click(sender As Object, e As RoutedEventArgs)
        AppCore.SaveWorkspace()
    End Sub

    Private Sub tlbr_Workspace_Delete_Click(sender As Object, e As RoutedEventArgs)
        AppCore.DeleteWorkspace()
    End Sub

    ' -- Keyboard shortcuts ---------------------------------------------------

    Private Sub Cmd_New_Executed(sender As Object, e As ExecutedRoutedEventArgs)
        AppCore.NewProject()
    End Sub

    Private Sub Cmd_Open_Executed(sender As Object, e As ExecutedRoutedEventArgs)
        AppCore.OpenProject()
    End Sub

    Private Sub Cmd_Save_Executed(sender As Object, e As ExecutedRoutedEventArgs)
        AppCore.SaveProject()
    End Sub

    Private Sub Cmd_SaveAs_Executed(sender As Object, e As ExecutedRoutedEventArgs)
        AppCore.SaveProjectAs()
    End Sub

    Private Sub Cmd_Help_Executed(sender As Object, e As ExecutedRoutedEventArgs)
        AppCore.ShowHelp()
    End Sub

End Class
