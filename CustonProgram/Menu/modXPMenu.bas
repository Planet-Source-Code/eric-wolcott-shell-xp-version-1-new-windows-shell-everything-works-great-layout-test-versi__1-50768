Attribute VB_Name = "modXPMenu"
Public XPSaveValue
Public Desktop As frmDesktop

Sub HandleClick(menuName As String, itemNum As Integer, strItemText As String)
    'MsgBox "Menu Name: " & menuName & vbCrLf & _
           "Item Number: " & itemNum & vbCrLf & _
           "Item Text: " & strItemText & vbCrLf & _
           "SaveValue: " & XPSaveValue
    Set Desktop = frmDesktop
    Select Case menuName
    Case "Icon"
                Select Case itemNum
                Case 1
                MsgBox "Opening " & frmDesktop.lblCaption(XPSaveValue).Tag
                Case 2
                Case 3
                Case 4
                Case 5
                Case 6
                'Load32Icon path, CLng(icon), imgIcon(imgIcon.UBound), Me ' - 1), Me
                Dim FRM2 As frmProperties
                Set FRM2 = frmProperties
                FRM2.LoadProps frmDesktop.lblCaption(XPSaveValue).Tag
                Case 7
                ExecuteShortcut2 frmDesktop.lblCaption(index).Tag
                End Select
    Case "Desktop"
                Select Case itemNum
                Case 1
                frmCreateShortcut.Show
                Case 2
                Case 3
                Desktop.LoadDesktop
                Case 4
                End Select
    Case "Menu"
                Select Case itemNum
                Case 1
                Case 2
                frmConsole.Visible = True
                Case 3
                frmTaskbar.Show
                Case 4
                frmProcesses.Show
                Case 5
                frmBandwidth.Show
                Case 6
                'End
                End Select
    End Select
End Sub


