VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBar 
   BorderStyle     =   0  'None
   Caption         =   "Taskbar"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6375
   Icon            =   "frmBar.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   3300
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   1605
      ScaleHeight     =   285
      ScaleWidth      =   300
      TabIndex        =   4
      Top             =   2715
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   3870
      ScaleHeight     =   450
      ScaleWidth      =   2355
      TabIndex        =   0
      Top             =   555
      Width           =   2355
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   105
         Width           =   1935
      End
      Begin VB.Image Image9 
         Height          =   450
         Left            =   1245
         Picture         =   "frmBar.frx":1042
         Top             =   0
         Width           =   75
      End
      Begin VB.Image Image8 
         Height          =   450
         Left            =   300
         Picture         =   "frmBar.frx":1264
         Top             =   0
         Width           =   75
      End
      Begin VB.Image Image10 
         Height          =   450
         Left            =   105
         Picture         =   "frmBar.frx":1486
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2220
      End
   End
   Begin Project1.API API 
      Height          =   480
      Left            =   540
      TabIndex        =   2
      Top             =   1980
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   1980
      Top             =   0
   End
   Begin VB.PictureBox Picture2 
      Height          =   330
      Left            =   270
      ScaleHeight     =   270
      ScaleWidth      =   1170
      TabIndex        =   3
      Top             =   1275
      Width           =   1230
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   210
      Top             =   2490
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
            Picture         =   "frmBar.frx":1540
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBar.frx":18DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBar.frx":1C74
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBar.frx":200E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBar.frx":206C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBar.frx":20CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBar.frx":2128
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image5 
      Height          =   450
      Left            =   2385
      Picture         =   "frmBar.frx":2186
      Top             =   2055
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Image Image4 
      Height          =   450
      Left            =   2280
      Picture         =   "frmBar.frx":6818
      Top             =   1440
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   1980
      TabIndex        =   5
      Top             =   2745
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   0
      Picture         =   "frmBar.frx":AEAA
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   75
      Left            =   15
      Picture         =   "frmBar.frx":BEEC
      Stretch         =   -1  'True
      Top             =   495
      Width           =   6225
   End
   Begin VB.Image Image11 
      Height          =   450
      Left            =   15
      Picture         =   "frmBar.frx":BFA6
      Stretch         =   -1  'True
      Top             =   570
      Width           =   6225
   End
   Begin VB.Image Image3 
      Height          =   450
      Index           =   0
      Left            =   1575
      Picture         =   "frmBar.frx":C060
      Top             =   2640
      Width           =   3000
   End
End
Attribute VB_Name = "frmBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public XPMenu As New clsXPMenu
Public XPMenu2 As New clsXPMenu
Public XPM_EFNet As New clsXPMenu
Public XPM_DALNet As New clsXPMenu

Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Sub Form_Load()
API.TaskBarHide
WindowPos Me, 1
Image11.Top = Image1.Height
Image11.Left = 0
Image1.Top = 0
Image1.Left = 0
Me.Width = Screen.Width
Image11.Width = Me.Width
Me.Height = Image11.Height + Image1.Height
Me.Top = Screen.Height - Me.Height
Me.Left = 0
Image1.Width = Me.Width
Image8.Left = 0
Image8.Top = 0
Image9.Top = 0
Image9.Left = Picture1.Width - Image9.Width
Image10.Width = Picture1.Width
Image10.Top = 0
Image10.Left = 0
Picture1.Top = Me.Height - Picture1.Height
Picture1.Left = Me.Width - Picture1.Width - 50
Image3(0).Top = Image11.Top
Image3(0).Left = 0 - (Image3(0).Width / 2)
Picture3(0).Top = Image3(0).Top + 3
Picture3(0).Left = Image3(0).Left + 3
End Sub

Private Sub Image2_Click()
Addbutton Me
'API.TaskBarShow
'Unload Me
End Sub

Private Sub Image3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3(Index).Picture = Image4.Picture
End Sub

Private Sub Image3_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3(Index).Picture = Image5.Picture
ShowWindow Image3(Index).Tag, 5
RemoveButton Image3(Index).Tag
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3(Index).Picture = Image4.Picture
End Sub

Private Sub Label1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
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
                
            XPMenu.ShowMenu pos.X, pos.Y
    Else
            ShowWindow Image3(Index).Tag, 5
            Image3(Index).Picture = Image5.Picture
            RemoveButton Image3(Index).Tag
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

Function Addbutton(frm27 As Form) As Long
Load Image3(Image3.UBound + 1)
Load Label1(Label1.UBound + 1)
Load Picture3(Picture3.UBound + 1)
Addbutton = Image3.UBound
        With Image3(Image3.UBound)
                .Top = Image3(Image3.UBound - 1).Top
                .Left = Image3(Image3.UBound - 1).Left + Image3(Image3.UBound - 1).Width
                .Visible = True
                .Tag = frm27.hwnd
                .ZOrder 0
        End With
        With Picture3(Picture3.UBound)
                .Top = Image3(Image3.UBound).Top + 75
                .Left = Image3(Image3.UBound).Left + 50
                .Picture = frm27.icon
                .AutoRedraw = True
                .Visible = True
                .ZOrder 0
        End With
        With Label1(Label1.UBound)
                .Caption = frm27.Caption
                .Top = Image3(Image3.UBound).Top + 100
                .Left = Picture3(Picture3.UBound).Left + Picture3(Picture3.UBound).Width + 75
                .Visible = True
                .ZOrder 0
        End With
End Function

Function RemoveButton(hwnd)
Dim FoundWindow
FoundWindow = 0
For X = 1 To Image3.UBound
If Image3(X).Tag = hwnd Then
FoundWindow = X
Exit For
End If
Next
If FoundWindow <> 0 Then
        Do While FoundWindow < Image3.UBound
        With Image3(FoundWindow)
                '.Top = Image3(FoundWindow + 1).Top
                '.Left = Image3(FoundWindow + 1).Left
                .Visible = True
                .Tag = Image3(FoundWindow + 1).Tag
                .ZOrder 0
        End With
        With Picture3(FoundWindow)
                '.Top = Picture3(FoundWindow + 1).Top
                '.Left = Picture3(FoundWindow + 1).Left
                .Picture = Picture3(FoundWindow + 1).Picture
                .AutoRedraw = True
                .Visible = True
                .ZOrder 0
        End With
        With Label1(FoundWindow)
                .Caption = Label1(FoundWindow + 1).Caption
                '.Top = Label1(FoundWindow + 1).Top
                '.Left = Label1(FoundWindow + 1).Left
                .Visible = True
                .ZOrder 0
        End With
        FoundWindow = FoundWindow + 1
        Loop
        Unload Image3(Image3.UBound)
        Unload Label1(Label1.UBound)
        Unload Picture3(Picture3.UBound)
Else
MsgBox "Error Removeing Button!"
End If
End Function
