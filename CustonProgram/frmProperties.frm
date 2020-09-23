VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmProperties 
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4635
   LinkTopic       =   "Form3"
   ScaleHeight     =   3180
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00F0F0F0&
      Height          =   2730
      Left            =   0
      TabIndex        =   0
      Top             =   450
      Width           =   4650
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   870
         TabIndex        =   13
         Text            =   "Text5"
         Top             =   1920
         Width           =   3660
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1755
         TabIndex        =   12
         Text            =   "Text4"
         Top             =   1560
         Width           =   765
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   870
         TabIndex        =   11
         Text            =   "Text3"
         Top             =   1560
         Width           =   795
      End
      Begin Project1.MacButton MacButton1 
         Height          =   330
         Left            =   3015
         TabIndex        =   10
         Top             =   2280
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   582
         BTYPE           =   4
         TX              =   "Save"
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         FCOL            =   0
      End
      Begin VB.PictureBox picTemp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   2400
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   9
         Top             =   210
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   870
         TabIndex        =   7
         Top             =   1200
         Width           =   3675
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   870
         TabIndex        =   6
         Top             =   810
         Width           =   3675
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   1455
         Top             =   300
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Icon Path : "
         Height          =   300
         Left            =   60
         TabIndex        =   14
         Top             =   1935
         Width           =   870
      End
      Begin VB.Image Picture1 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   3930
         Picture         =   "frmProperties.frx":0000
         Top             =   150
         Width           =   540
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "X/Y Pos : "
         Height          =   300
         Left            =   135
         TabIndex        =   8
         Top             =   1590
         Width           =   795
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Execute :"
         Height          =   240
         Left            =   75
         TabIndex        =   5
         Top             =   1200
         Width           =   750
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Caption :"
         Height          =   270
         Left            =   105
         TabIndex        =   4
         Top             =   810
         Width           =   720
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Icon :"
         Height          =   255
         Left            =   3405
         TabIndex        =   3
         Top             =   180
         Width           =   810
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Loaction :"
         Height          =   585
         Left            =   165
         TabIndex        =   2
         Top             =   195
         Width           =   3225
      End
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   4335
      Picture         =   "frmProperties.frx":0414
      Top             =   90
      Width           =   240
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Properties"
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
      Left            =   60
      TabIndex        =   1
      Top             =   90
      Width           =   3645
   End
   Begin VB.Image frmProperties 
      Height          =   450
      Left            =   0
      Picture         =   "frmProperties.frx":1456
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4635
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private fso As New FileSystemObject
Public strName27 As String
Public path27 As String, icon27 As String, Marker27 As String
Dim Desktop As frmDesktop
Function LoadProps(File As String)
Dim ff As Long
strName27 = File
ff = FreeFile
                Open File For Input As #ff
                Line Input #ff, path27
                Line Input #ff, icon27
                Line Input #ff, Marker27
                Close #ff
                
                        Dim x3, y3
                        x3 = Left(Marker27, InStr(1, Marker27, ",") - 1)
                        y3 = Right(Marker27, InStr(1, Marker27, ",") - 1)
                Text3.Text = x3
                Text4.Text = y3
                Label1.Caption = "Location : " & File
                Text2.Text = path27
                Picture1.Tag = icon27
                Text5.Text = icon27
                If UCase(Left(icon27, 4)) <> "APP," Then
                    icon27 = Replace(LCase(icon27), "%root%", App.path) 'ERoot)
                    Load32Icon icon27, 0, Picture1, Me '- 1), Me
                Else
                    icon27 = Right(icon27, Len(icon27) - InStr(1, icon27, ","))
                    Load32Icon path27, CLng(icon27), Picture1, Me ' - 1), Me
                End If
Me.Visible = True
End Function

Private Sub frmProperties_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hwnd, &HA1, 2, 0&
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hwnd, &HA1, 2, 0&
End Sub

Private Sub MacButton1_Click()
    With fso
        Set strm = .CreateTextFile(strName27, True)
        strm.Write (path27 & vbNewLine & Text5.Text & vbNewLine & Text3.Text & "," & Text4.Text)
    End With
Set Desktop = frmDesktop
Desktop.LoadDesktop
Unload Me
    Exit Sub
a:
MsgBox "an error has occurd"
Set Desktop = frmDesktop
Desktop.LoadDesktop
Unload Me
End Sub

Private Sub Picture1_Click()
CD1.Filter = ""
CD1.ShowOpen
If CD1.FileName = "" Then Exit Sub
Picture1.Tag = CD1.FileName
Text5.Text = Picture1.Tag
Load32Icon CD1.FileName, 0, Picture1, Me
End Sub
