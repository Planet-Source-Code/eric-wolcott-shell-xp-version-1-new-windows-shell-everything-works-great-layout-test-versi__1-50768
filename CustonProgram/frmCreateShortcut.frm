VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCreateShortcut 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "Create Shortcut"
   ClientHeight    =   3885
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5520
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCreateShortcut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.MacButton MacButton2 
      Height          =   240
      Left            =   4860
      TabIndex        =   6
      Top             =   690
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   423
      BTYPE           =   4
      TX              =   "..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   3
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      FCOL            =   0
   End
   Begin Project1.MacButton MacButton1 
      Height          =   435
      Left            =   3045
      TabIndex        =   5
      Top             =   2070
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   767
      BTYPE           =   4
      TX              =   "Create"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   3
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      FCOL            =   0
   End
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   4860
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   990
      Visible         =   0   'False
      Width           =   480
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2325
      Top             =   2850
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1215
      TabIndex        =   3
      Top             =   1275
      Width           =   3615
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   1215
      TabIndex        =   1
      Top             =   675
      Width           =   3615
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   5010
      Picture         =   "frmCreateShortcut.frx":0E42
      Top             =   75
      Width           =   240
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Create Shortcut"
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
      Left            =   1200
      TabIndex        =   8
      Top             =   60
      Width           =   3645
   End
   Begin VB.Image Image2 
      Height          =   450
      Left            =   1125
      Picture         =   "frmCreateShortcut.frx":1E84
      Stretch         =   -1  'True
      Top             =   -15
      Width           =   4245
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " Click To Change Icon"
      Height          =   255
      Left            =   1770
      TabIndex        =   7
      Top             =   1635
      Width           =   3315
   End
   Begin VB.Shape Shape1 
      Height          =   2205
      Left            =   1125
      Top             =   405
      Width           =   4245
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   255
      Left            =   1215
      TabIndex        =   2
      Top             =   1035
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Path:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1215
      TabIndex        =   0
      Top             =   435
      Width           =   1935
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Left            =   1215
      Picture         =   "frmCreateShortcut.frx":1F3E
      Top             =   1635
      Width           =   540
   End
End
Attribute VB_Name = "frmCreateShortcut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SaveTo As String
Public Desktop As frmDesktop

Private Sub Form_Load()
DoEvents
FadeIn Me, 255, 255
End Sub

Private Sub Image1_Click()

CD1.Filter = ""
CD1.ShowOpen

If CD1.FileName = "" Then Exit Sub

Image1.Tag = CD1.FileName
Load32Icon CD1.FileName, 0, Image1, Me

End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hwnd, &HA1, 2, 0&
End Sub

Private Sub Image3_Click()
Unload Me
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hwnd, &HA1, 2, 0&
End Sub

Private Sub MacButton1_Click()
Dim ff As Long
ff = FreeFile
Open App.path & "\Shortcuts\" & txtName & ".tmb" For Output As #ff
    Print #ff, txtFile
    Print #ff, Image1.Tag
    Print #ff, "PreSet,PreSet"
    DoEvents
Close #ff
Set Desktop = frmDesktop
Desktop.LoadDesktop
Unload Me
End Sub

Private Sub MacButton2_Click()
On Error Resume Next
CD1.Filter = "Applications|*.exe|All Files|*.*"
CD1.ShowOpen
If CD1.FileName = "" Then Exit Sub
If txtName = "" Then
    txtFile = CD1.FileName
    Dim p As Long
    p = Len(txtFile)
    Do Until Mid(txtFile, p, 1) = "\" Or p = 0
        p = p - 1
    Loop
    txtName = Right(txtFile, Len(textfile) - p)
    Load32Icon txtFile, 0, Image1, Me
    iamge1.Tag = txtFile
End If
End Sub

Private Sub txtFile_Change()
If Right(txtFile, 1) = "\" Then
    Load32Icon App.path & "\icon\folder.ico", 0, Image1, Me
    Image1.Tag = App.path & "\icon\folder.ico"
End If
End Sub

