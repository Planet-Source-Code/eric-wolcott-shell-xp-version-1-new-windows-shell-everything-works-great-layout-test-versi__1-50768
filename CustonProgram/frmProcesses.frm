VERSION 5.00
Begin VB.Form frmProcesses 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   1290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4260
      Picture         =   "frmProcesses.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   105
      Width           =   285
   End
   Begin Project1.CpuUsageControl CpuUsageControl1 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   90
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   1931
      BarColor        =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   90
      Left            =   -30
      Picture         =   "frmProcesses.frx":1042
      Stretch         =   -1  'True
      Top             =   1185
      Width           =   4650
   End
   Begin VB.Image Image7 
      Height          =   90
      Left            =   0
      Picture         =   "frmProcesses.frx":10FC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4650
   End
End
Attribute VB_Name = "frmProcesses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CpuUsageControl1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
ReleaseCapture
SendMessage Me.hwnd, &HA1, 2, 0&
End Sub

Private Sub Form_Load()
Me.Visible = False
WindowPos Me, 1
Me.Height = CpuUsageControl1.Height + (Image7.Height * 2)
Me.Top = frmBar.Top - Me.Height + Image7.Height
Me.Left = Screen.Width - Me.Width
Me.Visible = True
If Animation = True Then
FadeIn Me, 0, 255
End If
CpuUsageControl1.Enabled = True
CpuUsageControl1.StartMonitor
End Sub

Private Sub Picture1_Click()
If Animation = True Then
FadeOut Me, 255, 0
End If
Unload Me
End Sub
