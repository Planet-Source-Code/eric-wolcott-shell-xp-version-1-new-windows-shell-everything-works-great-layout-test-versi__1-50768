VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   2040
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   75
      Top             =   1515
   End
   Begin VB.TextBox txtConsole 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1410
      Left            =   45
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   6600
   End
   Begin VB.Timer Timer1 
      Left            =   5985
      Top             =   1020
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32" () As Long
Public Counter


Private Sub Form_Load()
Animation = True
Me.Width = Screen.Width
Me.Height = Screen.Height
WindowPos Me, 1
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
    Call keybd_event(&H5B, 0, 0, 0)
    Call keybd_event(&H4D, 0, 0, 0)
    Call keybd_event(&H5B, 0, &H2, 0)
FadeIn Me, 225, 255


Dim i As Integer
Dim M As Boolean
Dim t As Integer
Dim tick As Long, Lo As Long
Dim ttick As Long

Me.Show
DoEvents

ttick = GetTickCount
AddLine "___________________________________________________"
AddLine "ShellXP Version 1, loading, please wait", 12
AddLine "Current version is " & App.Major & "." & App.Minor & "." & App.Revision
AddLine "Written by Zach Szafran"
AddLine "___________________________________________________"
AddLine ""
AddLine "Loading Desktop..."
frmDesktop.Show
Me.Show
AddLine "Loading Attachments:"
AddLine "Loading Bar..."
frmBar.Show
AddLine "|Taskbar|"
DoEvents
frmTaskbar.Show
'AddLine "|Processor|"
'frmProcesses.Show
AddLine "|SystemTray|"
AddLine "Loading Core..."
AddLine "Reading Modules Config File..."
'LoadModuleConfig
AddLine ""
AddLine ""
AddLine "Loaded ShellXP in " & (GetTickCount - ttick) & "ms"
AddLine ""
AddLine ""
AddLine "Starting Countdown....."
Counter = 0
AddLine "Time (" & Counter & ")"
Timer2.Enabled = True
mL = t
End Sub

Public Function AddLine(txt As String, Optional FontSize As Integer)
txtConsole.Text = txtConsole.Text & txt & vbCrLf
txtConsole.SelStart = Len(txtConsole.Text)
frmConsole.txtConsole.Text = txtConsole.Text
frmConsole.txtConsole.SelStart = Len(frmConsole.txtConsole.Text)
DoEvents
End Function

Public Function AddLine2(txt As String, Start As Integer, Legnth As Integer)
txtConsole.SelStart = Start
txtConsole.SelLength = Legnth
txtConsole.SelText = txt
txtConsole.SelStart = Len(txtConsole.Text)
End Function

Private Sub Image2_Click()
Me.Visible = False
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hwnd, &HA1, 2, 0&
End Sub

Private Sub Timer2_Timer()
If Counter < 0 Then
Me.Visible = False
If Animation = True Then
Load Form3
End If
Timer2.Enabled = False
End If

AddLine2 Counter & ")", Len(txtConsole.Text) - 2 - Len(vbCrLf), 2
DoEvents
Counter = Counter - 1
End Sub

Private Sub Form_Resize()
txtConsole.Height = Me.Height
txtConsole.Top = 0
txtConsole.Left = o
txtConsole.Width = Me.Width
TransparentForm Me
End Sub

