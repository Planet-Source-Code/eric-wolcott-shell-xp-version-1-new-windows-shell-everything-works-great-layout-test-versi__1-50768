VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   2910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6570
   LinkTopic       =   "Form3"
   ScaleHeight     =   2910
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.Display Display1 
      Height          =   330
      Left            =   30
      TabIndex        =   3
      Top             =   465
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   582
      Characters      =   27
      Color           =   65535
      Caption         =   "Display"
      ScrollRate      =   1
      Scroll          =   0
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":0000
      ForeColor       =   &H00FFFFFF&
      Height          =   1650
      Left            =   45
      TabIndex        =   4
      Top             =   1140
      Width           =   6375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Created By Zach Szafran"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   60
      TabIndex        =   2
      Top             =   915
      Width           =   6375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SHELLXP version 1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   150
      TabIndex        =   1
      Top             =   525
      Width           =   6285
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   2475
      Left            =   0
      Top             =   435
      Width           =   6570
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   30
      TabIndex        =   0
      Top             =   75
      Width           =   3645
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   6255
      Picture         =   "frmAbout.frx":01D4
      Top             =   105
      Width           =   240
   End
   Begin VB.Image Image7 
      Height          =   450
      Left            =   0
      Picture         =   "frmAbout.frx":1216
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6570
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   45
      Picture         =   "frmAbout.frx":12D0
      Top             =   2325
      Width           =   1860
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
If Animation = True Then
FadeIn Me, 0, 255
End If
End Sub

Private Sub Image3_Click()
If Animation = True Then
FadeOut Me, 255, 0
End If
Unload Me
End Sub

Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ReleaseCapture
SendMessage Me.hwnd, &HA1, 2, 0&
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ReleaseCapture
SendMessage Me.hwnd, &HA1, 2, 0&
End Sub
