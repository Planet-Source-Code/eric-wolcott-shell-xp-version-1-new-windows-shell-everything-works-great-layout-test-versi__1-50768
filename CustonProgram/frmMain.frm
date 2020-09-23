VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "frmMain"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3210
      TabIndex        =   3
      Top             =   2295
      Width           =   1215
   End
   Begin VB.Timer tmrMoveIcon 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3240
      Top             =   480
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   690
      TabIndex        =   2
      Top             =   1485
      Width           =   2160
   End
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   2685
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   1260
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Timer tmrDesktop 
      Interval        =   5
      Left            =   4140
      Top             =   60
   End
   Begin MSWinsockLib.Winsock wsckModule 
      Left            =   3720
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   15
      LocalPort       =   16
   End
   Begin VB.Shape shpMove 
      BorderStyle     =   3  'Dot
      Height          =   915
      Left            =   1560
      Top             =   900
      Visible         =   0   'False
      Width           =   735
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
      Left            =   180
      TabIndex        =   0
      Top             =   660
      Width           =   630
   End
   Begin VB.Image imgIcon 
      Height          =   465
      Index           =   0
      Left            =   240
      Top             =   180
      Width           =   480
   End
   Begin VB.Image imgDesktop 
      Height          =   435
      Left            =   0
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const mName = "Desktop"
Public ERoot As String
Private Whwnd As Long

Private mx As Long, my As Long

Private Sub Command1_Click()
Dim x
For x = 0 To imgIcon.UBound - 1
ChangeXYIcon lblCaption(x).Tag, 175, x * (imgIcon(0).Height + lblCaption(0).Height) + 750
LoadDesktop
Next
End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
Me.Height = Screen.Height
Me.Width = Screen.Width

If Command$ <> "-nocore" Then

    wsckModule.SendData "CORE,LOADED"
    
Else

    ERoot = "c:\vb\eshell beta2"
    tmrDie.Enabled = False
    ReloadDesktopBG
    LoadDesktop
    
End If

Whwnd = Me.hWnd
DoEvents
SetDesktop Whwnd, Me


End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 2 Then Me.PopupMenu frmMenu.mnuDesk, , x, y

End Sub

Private Sub imgDesktop_Click()
LoadDesktop
End Sub

Private Sub imgDesktop_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Call Form_MouseUp(Button, Shift, imgDesktop.Left + x, imgDesktop.Top + y)
End Sub

Private Sub imgIcon_DblClick(Index As Integer)

wsckModule.SendData "CORE,LOADESL," & lblCaption(Index).Tag

End Sub

Private Sub imgIcon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 1 Then

    shpMove.Visible = True

    shpMove.Left = imgIcon(Index).Left
    If lblCaption(Index).Left < shpMove.Left Then shpMove.Left = lblCaption(Index).Left
    
    shpMove.Top = imgIcon(Index).Top
    
    shpMove.Height = lblCaption(Index).Top + lblCaption(Index).Height - imgIcon(Index).Top
    
    If lblCaption(Index).Width > imgIcon(Index).Width Then
    
        shpMove.Width = lblCaption(Index).Width
    
    Else
    
        shpMove.Width = imgIcon(Index).Width
    
    End If
    
    mx = imgIcon(Index).Left + x - shpMove.Left
    my = imgIcon(Index).Top + y - shpMove.Top

    tmrMoveIcon.Enabled = True
    
ElseIf Button = 2 Then

    frmMenu.cIcon = Index
    Me.PopupMenu frmMenu.mnuIcon, , imgIcon(Index).Left + x, imgIcon(Index).Top + y

End If

End Sub

Private Sub imgIcon_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

tmrMoveIcon.Enabled = False

shpMove.Visible = False
lblCaption(Index).Top = shpMove.Top + shpMove.Height - 210
imgIcon(Index).Top = shpMove.Top
imgIcon(Index).Left = shpMove.Left + shpMove.Width / 2 - imgIcon(Index).Width / 2
lblCaption(Index).Left = shpMove.Left + shpMove.Width / 2 - lblCaption(Index).Width / 2

ChangeXYIcon lblCaption(Index).Tag, imgIcon(Index).Left, imgIcon(Index).Top

End Sub

Private Sub lblCaption_DblClick(Index As Integer)

wsckModule.SendData "CORE,LOADESL," & lblCaption(Index).Tag

End Sub

Private Sub tmrDesktop_Timer()

If Whwnd = GetActiveWindow Then

SetDesktop Whwnd, Me

End If

End Sub

Private Sub tmrDie_Timer()
End
End Sub

Private Sub tmrMoveIcon_Timer()

Dim x As Long, y As Long

shpMove.Left = GetX * 15 - mx
shpMove.Top = GetY * 15 - my


End Sub

Public Function ReloadDesktopBG()
ERoot = App.path
Me.BackColor = ReadValue("desktop", "bgcol", ERoot & "\eshell.cfg", "&H3A6EA5&")
    
If ReadValue("desktop", "bg", ERoot & "\eshell.cfg", "") <> "" Then
    
    imgDesktop.Picture = LoadPicture(ReadValue("desktop", "bg", ERoot & "\eshell.cfg", ""))
    
    If ReadValue("desktop", "stretch", ERoot & "\eshell.cfg", False) = False Then
    
        imgDesktop.Stretch = False
        imgDesktop.Left = Me.Width / 2 - imgDesktop.Width / 2
        imgDesktop.Top = Me.Height / 2 - imgDesktop.Height / 2
    
    Else
    
        imgDesktop.Stretch = True
        Me.BackColor = ReadValue("desktop", "bgcol", ERoot & "\eshell.cfg", "&H3A6EA5&")
        
        imgDesktop.Left = 0
        imgDesktop.Top = 0
        imgDesktop.Height = Me.Height
        imgDesktop.Width = Me.Width
    
    End If

End If

End Function
