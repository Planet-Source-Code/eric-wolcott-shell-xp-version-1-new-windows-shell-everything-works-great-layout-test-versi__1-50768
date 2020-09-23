Attribute VB_Name = "modTrans"
Option Explicit
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Const WS_EX_TRANSPARENT = &H20&
Public Const LWA_ALPHA = &H2&
Private Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Sub TransparentForm(frm As Form)
    On Error Resume Next
frm.ScaleMode = vbPixels
Const RGN_DIFF = 4
Const RGN_OR = 2
Dim outer_rgn As Long
Dim inner_rgn As Long
Dim wid As Single
Dim hgt As Single
Dim border_width As Single
Dim title_height As Single
Dim ctl_left As Single
Dim ctl_top As Single
Dim ctl_right As Single
Dim ctl_bottom As Single
Dim control_rgn As Long
Dim combined_rgn As Long
Dim ctl As Control

If frm.WindowState = vbMinimized Then Exit Sub
wid = frm.ScaleX(frm.Width, vbTwips, vbPixels)
hgt = frm.ScaleY(frm.Height, vbTwips, vbPixels)
outer_rgn = CreateRectRgn(0, 0, wid, hgt)
border_width = (wid - frm.ScaleWidth) / 2
title_height = hgt - border_width - frm.ScaleHeight
inner_rgn = CreateRectRgn(border_width, title_height, wid - border_width, _
hgt - border_width)
combined_rgn = CreateRectRgn(0, 0, 0, 0)
CombineRgn combined_rgn, outer_rgn, inner_rgn, RGN_DIFF
DoEvents
For Each ctl In frm.Controls
If ctl.Container Is frm Then
ctl_left = frm.ScaleX(ctl.Left, frm.ScaleMode, vbPixels) _
+ border_width
ctl_top = frm.ScaleX(ctl.Top, frm.ScaleMode, vbPixels) + title_height
ctl_right = frm.ScaleX(ctl.Width, frm.ScaleMode, vbPixels) + ctl_left
ctl_bottom = frm.ScaleX(ctl.Height, frm.ScaleMode, vbPixels) + ctl_top
control_rgn = CreateRectRgn(ctl_left, ctl_top, ctl_right, ctl_bottom)
CombineRgn combined_rgn, combined_rgn, control_rgn, RGN_OR
End If
DoEvents
Next ctl
SetWindowRgn frm.hwnd, combined_rgn, True
End Sub

Public Sub FadeIn(frm As Form, intStartValue As Integer, intEndValue As Integer)
frm.Visible = False
TransparentForm frm
DoEvents
frm.Visible = True
DoEvents
Dim NormalWindowStyle As Long
NormalWindowStyle = GetWindowLong(frm.hwnd, GWL_EXSTYLE)
SetWindowLong frm.hwnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED
DoEvents
Dim X
For X = intStartValue To intEndValue
SetLayeredWindowAttributes frm.hwnd, 0, X, LWA_ALPHA
frm.Refresh
DoEvents
Next
End Sub
Public Sub FadeOut(frm As Form, intStartValue As Integer, intEndValue As Integer)
Dim NormalWindowStyle As Long
NormalWindowStyle = GetWindowLong(frm.hwnd, GWL_EXSTYLE)
SetWindowLong frm.hwnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED
DoEvents
Dim X
For X = intEndValue To intStartValue
SetLayeredWindowAttributes frm.hwnd, 0, 255 - X, LWA_ALPHA
frm.Refresh
DoEvents
Next
End Sub

Public Sub FadeIn2(frm As Form, intStartValue As Integer, intEndValue As Integer)
frm.Visible = True
DoEvents
Dim NormalWindowStyle As Long
NormalWindowStyle = GetWindowLong(frm.hwnd, GWL_EXSTYLE)
SetWindowLong frm.hwnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED
DoEvents
Dim X
For X = intStartValue To intEndValue
SetLayeredWindowAttributes frm.hwnd, 0, X, LWA_ALPHA
frm.Refresh
DoEvents
Next
End Sub


