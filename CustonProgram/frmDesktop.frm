VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDesktop 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6270
   LinkTopic       =   "Form2"
   ScaleHeight     =   4530
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.API API 
      Height          =   480
      Left            =   3165
      TabIndex        =   4
      Top             =   3495
      Visible         =   0   'False
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   2490
      Top             =   765
   End
   Begin VB.Timer tmrMoveIcon 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1980
      Top             =   765
   End
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3060
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   765
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   2055
      TabIndex        =   0
      Top             =   1860
      Visible         =   0   'False
      Width           =   2895
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2460
      Top             =   1230
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesktop.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesktop.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesktop.frx":0734
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesktop.frx":0ACE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesktop.frx":1B20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesktop.frx":2B72
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesktop.frx":3BC4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00E0E0E0&
      Height          =   330
      Index           =   3
      Left            =   1500
      Shape           =   3  'Circle
      Top             =   3300
      Width           =   330
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00E0E0E0&
      Height          =   330
      Index           =   2
      Left            =   855
      Shape           =   3  'Circle
      Top             =   3330
      Width           =   540
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00E0E0E0&
      Height          =   330
      Index           =   1
      Left            =   1395
      Shape           =   3  'Circle
      Top             =   2610
      Width           =   330
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00E0E0E0&
      Height          =   330
      Index           =   0
      Left            =   840
      Shape           =   3  'Circle
      Top             =   2610
      Width           =   330
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   105
      Picture         =   "frmDesktop.frx":4C16
      Top             =   105
      Width           =   240
   End
   Begin VB.Image Image7 
      Height          =   450
      Left            =   0
      Picture         =   "frmDesktop.frx":5C58
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6225
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1695
      TabIndex        =   3
      Top             =   3390
      Width           =   1170
   End
   Begin VB.Image Image6 
      Height          =   720
      Left            =   1860
      Picture         =   "frmDesktop.frx":5D12
      Top             =   3135
      Width           =   720
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   3075
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Shape shpMove 
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      Height          =   660
      Left            =   1110
      Top             =   1770
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image imgIcon 
      Height          =   465
      Index           =   0
      Left            =   210
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[Caption]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   150
      TabIndex        =   1
      Top             =   990
      Visible         =   0   'False
      Width           =   630
   End
End
Attribute VB_Name = "frmDesktop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public XPMenu As New clsXPMenu
Public XPMenu2 As New clsXPMenu
Public XPM_EFNet As New clsXPMenu
Public XPM_DALNet As New clsXPMenu

Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Form As Form1
Private Whwnd As Long
Dim ShapePlace(5) As Boolean
Dim ShapeNumber As Integer
Public mX As Integer
Public mY As Integer

Private Sub Command1_Click()
LoadDesktop
End Sub

Private Sub Form_Load()
Set Form = Form1
Me.Top = 0
Me.Left = 0
Me.Width = Screen.Width
Me.Height = Screen.Height ' - 500

Form1.Top = Image7.Height
Form1.Left = 0
Form1.Height = Screen.Height - Image7.Height - 75
Form1.Width = Me.Width

Image6.Top = (Me.Height - Image6.Height) / 2
Image6.Left = (Me.Width - Image6.Width) / 2

Label1.Width = Image6.Width
Label1.Left = Image6.Left
Label1.Top = Image6.Top + (Image6.Height / 4)

Image7.Top = 0
Image7.Width = Me.Width
Image7.Left = 0

Shape1(1).Top = 3315
Shape1(1).Left = 4230
Shape1(2).Top = 3345
Shape1(2).Left = 7125
Shape1(3).Top = 5235
Shape1(3).Left = 4440
Shape1(0).Top = 5175
Shape1(0).Left = 7095
FadeIn2 Me, 255, 255
LoadDesktop
End Sub

Public Function LoadDesktop()
'On Error Resume Next
Dim i As Long
Dim pth As String
Dim ff As Long
Dim L As String
Dim p As Long
Dim Path As String, icon As String, Marker As String
Dim x As Long, y As Long

If imgIcon.UBound > 0 Then
    For i = 1 To imgIcon.UBound
        Unload imgIcon(i)
        Unload lblCaption(i)
    Next i
End If

pth = App.Path & "\Shortcuts\"
ff = FreeFile
File1.Path = "C:\"
File1.Path = pth
Form.AddLine "-Save Path:" & File1.Path
i = 0
ShapeNumber = 0
For i = 0 To File1.ListCount - 1
    If Right(File1.List(i), 4) = ".tmb" Then
        Load imgIcon(imgIcon.UBound + 1)
        Load lblCaption(lblCaption.UBound + 1)
        
        With imgIcon(imgIcon.UBound)
            .Visible = False
            .Top = imgIcon(ShapeNumber).Top + .Height + 210 + 180
            .ZOrder 0
            .ToolTipText = File1.List(i)
        End With
        With lblCaption(imgIcon.UBound)
            .Visible = False
            .Top = lblCaption(ShapeNumber).Top + imgIcon(ShapeNumber).Height + 210 + 180
            .ZOrder 0
        End With
        
        Form.AddLine "-Loading Shortcut: " & File1.List(i)
                ff = FreeFile
                Open pth & File1.List(i) For Input As #ff
                Line Input #ff, Path
                Line Input #ff, icon
                Line Input #ff, Marker
                Close #ff
                If UCase(Left(icon, 4)) <> "APP," Then
                    icon = Replace(LCase(icon), "%root%", App.Path) 'ERoot)
                    Load32Icon icon, 0, imgIcon(imgIcon.UBound), Me '- 1), Me
                    Form.AddLine "-Loading Shurtcut Icon: " & icon
                Else
                    icon = Right(icon, Len(icon) - InStr(1, icon, ","))
                    Load32Icon Path, CLng(icon), imgIcon(imgIcon.UBound), Me ' - 1), Me
                    Form.AddLine "-Loading Shurtcut Icon: " & icon
                End If
                Dim ONOff As Boolean
                ONOff = True
                If Marker <> "" Then
                With imgIcon(imgIcon.UBound)
                        Dim x3, y3
                        x3 = Left(Marker, InStr(1, Marker, ",") - 1)
                        y3 = Right(Marker, InStr(1, Marker, ",") - 1)
                        If InStr(1, x3, "PreSet") = False And InStr(1, x3, "Space") = False Then
                                .Top = x3
                                lblCaption(lblCaption.UBound).Top = .Top + .Height
                                ONOff = False
                        ElseIf InStr(1, x3, "Space") Then
                                .Top = Shape1(Right(x3, InStr(1, x3, "-") - 5)).Top
                                lblCaption(lblCaption.UBound).Top = .Top + .Height
                                ONOff = False
                        End If
                        
                        If InStr(1, y3, "PreSet") = False And InStr(1, y3, "Space") = False Then
                                .Left = y3
                                lblCaption(lblCaption.UBound).Left = .Left + .Height
                                ONOff = False
                        ElseIf InStr(1, y3, "Space") Then
                                .Left = Shape1(Right(x3, InStr(1, x3, "-") - 5)).Left
                                lblCaption(lblCaption.UBound).Left = .Left + .Height
                                ONOff = False
                        End If
                End With
                End If
                If ONOff = True Then
                ShapeNumber = imgIcon.UBound
                End If
        DoEvents
       
        With lblCaption(lblCaption.UBound) '- 1)
        .Caption = Left(File1.List(i), Len(File1.List(i)) - 4)
        Form.AddLine "-Shortcut Loaded Sucessfully"
        .Tag = pth & File1.List(i)
rewidth:
        
        If .Width > 960 Then
            If Right(.Caption, 3) <> "..." Then .Caption = .Caption & "..."
            .Caption = Left(.Caption, Len(.Caption) - 4)
            .Caption = .Caption & "..."
            GoTo rewidth
        End If
        .Left = imgIcon(imgIcon.UBound).Left + (imgIcon(imgIcon.UBound).Width / 2) - (.Width / 2)
        End With
    End If
Next i
Dim e
For e = 1 To imgIcon.UBound
        imgIcon(e).Visible = True
        lblCaption(e).Visible = True
Next
Form.AddLine "-" & imgIcon.UBound & " Desktop Icons Loaded"
End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then
    XPMenu.Init "Desktop", ImageList1
    XPMenu.AddItem 0, "Create Shortcut", False, False, XPMenu2
    XPMenu.AddItem 0, "", False, True
    XPMenu.AddItem 1, "Line-Up Icons", False, False
    XPMenu.AddItem 2, "Counsole", False, False
    XPMenu.AddItem 2, "Task Panel", False, False
    XPMenu.AddItem 0, "", False, True
    XPMenu.AddItem 3, "Preferences", False, False
    XPMenu.AddItem 0, "", False, True
    XPMenu.AddItem 0, "Exit ShellXP", False, False
    
    Dim pos As POINTAPI
    GetCursorPos pos
        
    XPMenu.ShowMenu pos.x, pos.y
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'StartMenu.Hide
End Sub

Private Sub Form_Terminate()
API.TaskBarShow
End Sub

Private Sub Form_Unload(Cancel As Integer)
API.TaskBarShow
End Sub

Private Sub Image1_Click()
API.TaskBarShow
'FadeOut Me, 255, 0
End
End Sub

Private Sub Image6_Click()
    XPMenu.Init "Menu", ImageList1
    XPMenu.AddItem 0, "ShellXP v1.0 Menu", False, False
    XPMenu.AddItem 2, "Counsole", False, False
    XPMenu.AddItem 2, "Task Panel", False, False
    XPMenu.AddItem 2, "CPU Usage", False, False
    XPMenu.AddItem 2, "Bandwidth Usage", False, False
    XPMenu.AddItem 2, "Exit ShellXP", False, False
    
    Dim pos As POINTAPI
    GetCursorPos pos
    XPMenu.ShowMenu (Image6.Left - 600) / 15, (Image6.Top + 400) / 15
End Sub

Private Sub imgIcon_DblClick(index As Integer)
ExecuteShortcut lblCaption(index).Tag
End Sub

Private Sub imgIcon_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    shpMove.Top = imgIcon(index).Top
    shpMove.Left = imgIcon(index).Left
    shpMove.Width = imgIcon(index).Width
    shpMove.Height = imgIcon(index).Height
    shpMove.Visible = True
    mX = GetX - imgIcon(index).Left 'imgIcon(index).Left + GetX - shpMove.Left
    mY = GetY - imgIcon(index).Top 'imgIcon(index).Top + GetY - shpMove.Top
    tmrMoveIcon.Enabled = True
    Exit Sub
ElseIf Button = 2 Then
    XPSaveValue = index
    XPMenu.Init "Icon", ImageList1
    XPMenu.AddItem 4, "Open", False, False, XPMenu2
    XPMenu.AddItem 0, "", False, True
    XPMenu.AddItem 5, "Rename", False, False
    XPMenu.AddItem 6, "Delete", False, False
    XPMenu.AddItem 0, "", False, True
    XPMenu.AddItem 7, "Properties", False, False
    XPMenu.AddItem 0, "", False, False
    XPMenu.AddItem 6, "Resource Edit", False, False
    Dim pos As POINTAPI
    GetCursorPos pos
    XPMenu.ShowMenu pos.x, pos.y
    End If
End Sub

Private Sub imgIcon_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
tmrMoveIcon.Enabled = False
shpMove.Visible = False
imgIcon(index).Visible = False
lblCaption(index).Visible = False
lblCaption(index).Top = shpMove.Top + shpMove.Height - 210 + lblCaption(index).Height
imgIcon(index).Top = shpMove.Top
imgIcon(index).Left = shpMove.Left + shpMove.Width / 2 - imgIcon(index).Width / 2
lblCaption(index).Left = shpMove.Left + shpMove.Width / 2 - lblCaption(index).Width / 2
imgIcon(index).Visible = True
lblCaption(index).Visible = True
'ChangeXYIcon lblCaption(Index).Tag, imgIcon(Index).Left, imgIcon(Index).Top
End If
End Sub


Private Sub Label1_Click()
Call Image6_Click
End Sub

Private Sub Timer1_Timer()
If Whwnd = GetActiveWindow Then
SetDesktop Whwnd, Me
End If
End Sub


Private Sub Timer3_Timer()
Dim Hr, Mn, Sc, Dn
Hr = Hour(Now)
If Hr - 12 > 0 Then
Hr = Hr - 12
Dn = "PM"
Else
dm = "AM"
End If
Mn = Minute(Now)
If Len(Mn) = 1 Then
Mn = Mn & "0"
End If
Sc = Second(Now)
Label2.Caption = Hr & ":" & Mn & "." & Sc & " " & Dn
End Sub

Private Sub tmrMoveIcon_Timer()
shpMove.Visible = False
shpMove.Left = GetX - mX
shpMove.Top = GetY - mY
shpMove.Visible = True
End Sub
