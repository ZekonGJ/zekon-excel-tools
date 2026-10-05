Option Explicit
Private Sub Workbook_Open()
    InstallMenu
End Sub
Private Sub Workbook_AddinInstall()
    InstallMenu
End Sub
Private Sub Workbook_AddinUninstall()
    Auto_Close
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Auto_Close
End Sub
