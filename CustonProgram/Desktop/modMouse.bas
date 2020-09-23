Attribute VB_Name = "modMouse"
Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public P As POINTAPI
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Function GetX()
Dim P As POINTAPI
GetCursorPos P
GetX = P.X * 15 - mx
End Function

Public Function GetY()
Dim P As POINTAPI
GetCursorPos P
GetY = P.Y * 15 - my
End Function

