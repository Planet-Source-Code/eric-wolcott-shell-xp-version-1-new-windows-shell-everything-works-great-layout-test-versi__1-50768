VERSION 5.00
Begin VB.Form frmStartMenu 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   4845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6555
   LinkTopic       =   "Form3"
   ScaleHeight     =   4845
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   1980
      TabIndex        =   6
      Top             =   4455
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4140
      Left            =   435
      TabIndex        =   2
      Top             =   195
      Width           =   2580
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   0
         Left            =   60
         ScaleHeight     =   405
         ScaleWidth      =   2430
         TabIndex        =   5
         Top             =   3645
         Width           =   2460
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Run"
            Height          =   210
            Left            =   330
            TabIndex        =   7
            Top             =   105
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   0
            Left            =   45
            Picture         =   "frmStartMenu.frx":0000
            Top             =   90
            Width           =   240
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   45
         Picture         =   "frmStartMenu.frx":1042
         ScaleHeight     =   285
         ScaleWidth      =   300
         TabIndex        =   4
         Top             =   285
         Visible         =   0   'False
         Width           =   300
      End
      Begin Project1.chameleonButton chameleonButton2 
         Height          =   4140
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   7303
         BTYPE           =   4
         TX              =   ""
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         FCOL            =   0
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   360
      Picture         =   "frmStartMenu.frx":2084
      ScaleHeight     =   270
      ScaleWidth      =   645
      TabIndex        =   1
      Top             =   4485
      Width           =   645
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   0
      Top             =   975
   End
   Begin Project1.chameleonButton chameleonButton1 
      Height          =   510
      Left            =   15
      TabIndex        =   0
      Top             =   4335
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   900
      BTYPE           =   4
      TX              =   "Menu"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   3
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      FCOL            =   0
   End
End
Attribute VB_Name = "frmStartMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MenuVisible As Boolean
Private Whwnd As Long
Private Sub chameleonButton1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Frame1.Top = chameleonButton1.Top - Frame1.Height
Picture2.Top = Frame1.Top + 5
Picture2.Visible = True
MenuVisible = True
Frame1.Visible = True
TransparentForm Me
'Top = 4305
End Sub

Private Sub chameleonButton2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim f
For f = 0 To Picture1.UBound
If Picture1(f).BackColor <> &HF0F0F0 Then
Picture1(f).BackColor = &HF0F0F0
End If
Next
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Frame1.Top = Me.Height
Me.Height = Screen.Height
List1.Top = Me.Height
chameleonButton1.Top = Me.Height - chameleonButton1.Height
Picture3.Top = Me.Height - Picture3.Height - 100
MenuVisible = True
HideMenu
DoEvents
MakeTopMost Me.hwnd
WindowPos Me, 1
FadeIn Me, 200, 255
AddStartMenu
AddStartMenu
AddStartMenu
End Sub

Function HideMenu()
If MenuVisible = True Then
Frame1.Top = Me.Height
Picture2.Top = Me.Height - Picture2.Height
Picture2.Visible = False
TransparentForm Me
MenuVisible = False
End If
End Function

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hwnd, &HA1, 2, 0&
End Sub

Private Sub Picture1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Shell (InputBox("Shell"))
End Sub

Private Sub Picture1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1(Index).BackColor = &HFBEFE6
End Sub

Private Sub Picture2_Click()
HideMenu
End Sub

Private Sub Timer1_Timer()
If Whwnd = GetActiveWindow Then
SetDesktop Whwnd, Me
End If
End Sub

Function AddStartMenu()
Load Picture1(Picture1.UBound + 1)
Load Image1(Image1.UBound + 1)


Picture1(Picture1.UBound).Top = Picture1(Picture1.UBound - 1).Top - 5 - Picture1(Picture1.UBound).Height
Picture1(Picture1.UBound).Container = Frame1
Picture1(Picture1.UBound).Visible = True

Image1(Image1.UBound).Top = Picture1(Picture1.UBound).Top - 5
Image1(Image1.UBound).Container = Picture1(Picture1.UBound)
Image1(Image1.UBound).Visible = True
End Function
