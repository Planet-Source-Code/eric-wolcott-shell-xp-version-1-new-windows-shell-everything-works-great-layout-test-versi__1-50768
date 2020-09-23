Attribute VB_Name = "modExecuteIcon"
Public Sub ExecuteShortcut(Path As String)
        'Form.AddLine "-Execute Shortcut: " & Path
                ff = FreeFile
                Open Path For Input As #ff
                Line Input #ff, Path
                Line Input #ff, icon
                Line Input #ff, Marker
                Close #ff
If Left(Path, Len("[CONSOLE]")) = "[CONSOLE]" Then
Dim temp
temp = Right(Path, Len(Path) - Len("[CONSOLE] "))
ConsoleExecute temp
Else
ShellFile Path
End If
End Sub

Public Sub ExecuteShortcut2(Path As String)
                ff = FreeFile
                Open Path For Input As #ff
                Line Input #ff, Path2
                Line Input #ff, icon
                Line Input #ff, Marker
                Close #ff
End Sub

Public Sub ConsoleExecute(Tag)
Select Case UCase(Tag)
Case "ABOUT"
frmAbout.Show
End Select
End Sub

