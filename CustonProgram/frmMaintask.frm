VERSION 5.00
Begin VB.Form frmTaskbar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Tasklisting"
   ClientHeight    =   3480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3450
   Icon            =   "frmMaintask.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   3450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture6 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   105
      Picture         =   "frmMaintask.frx":0E42
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   12
      Top             =   285
      Width           =   240
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   105
      Picture         =   "frmMaintask.frx":1E84
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   11
      Top             =   45
      Width           =   255
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   0
      Left            =   480
      ScaleHeight     =   330
      ScaleWidth      =   300
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   300
   End
   Begin Project1.MacButton Command1 
      Height          =   510
      Index           =   0
      Left            =   405
      TabIndex        =   9
      Top             =   15
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   900
      BTYPE           =   4
      TX              =   ""
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
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   105
      TabIndex        =   7
      Top             =   165
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2010
      Top             =   1980
   End
   Begin VB.ListBox lstApps 
      Height          =   255
      Left            =   105
      TabIndex        =   3
      Top             =   165
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lstNames 
      Height          =   255
      Left            =   105
      TabIndex        =   2
      Top             =   150
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lstHwnd 
      Height          =   255
      Left            =   105
      TabIndex        =   1
      Top             =   165
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lstHwndNames 
      Height          =   255
      Left            =   90
      TabIndex        =   0
      Top             =   150
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox Picture1 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   4
      Top             =   0
      Width           =   0
   End
   Begin VB.PictureBox Picture2 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   5
      Top             =   0
      Width           =   0
   End
   Begin VB.PictureBox Picture3 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   6
      Top             =   0
      Width           =   0
   End
   Begin Project1.MacButton MacButton1 
      Height          =   3345
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   5900
      BTYPE           =   4
      TX              =   ""
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
      BCOL            =   8421504
      FCOL            =   0
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   105
      Picture         =   "frmMaintask.frx":2EC6
      Top             =   135
      Width           =   240
   End
End
Attribute VB_Name = "frmTaskbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public X2 As Integer, Y2 As Integer
Public Form As Form1
Dim Bar As frmBar
Function CheckApps()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Ceck For Lost Apps
Dim Missing As Boolean, AllMissing As Boolean
Dim MissingName As String, MissingID As Long
If lstHwnd.ListCount > 0 Then
AllMissing = False
For X2 = 0 To lstApps.ListCount - 1
Missing = False
        For Y2 = 0 To lstApps.ListCount - 1
        If lstApps.List(Y2) <> lstHwnd.List(X2) Then
        Missing = True
        Else
        AllMissing = True
        Missing = False
        Exit For
        End If
        Next
        If Missing = True And lstHwndNames.List(X2) <> "" Then
        List1.AddItem "L-" & lstHwndNames.List(X2)
        MissingName = lstHwndNames.List(X2)
        MissingID = lstHwnd.List(X2)
        Add_Remove_Button False, True, MissingID
        End If
Next
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Ceck For New Apps
Dim Found As Boolean, AllFound As Boolean
Dim FoundName As String, FoundID As Long
AllFound = False
For X2 = 0 To lstApps.ListCount - 1
Found = False
        For Y2 = 0 To lstHwnd.ListCount - 1
        If lstApps.List(X2) <> lstHwnd.List(Y2) Then
        Found = True
        Else
        AllFound = True
        Found = False
        Exit For
        End If
        Next
        If Found = True And lstNames.List(X2) <> "" Then
        List1.AddItem "F-" & lstNames.List(X2)
        FoundName = lstNames.List(X2)
        FoundID = lstApps.List(X2)
        Add_Remove_Button True, False, FoundID, CheckCaption(FoundName), FoundName
        End If
Next

If AllMissing = True Or AllFound = True Or lstHwnd.ListCount < 1 Then
lstHwnd.Clear
lstHwndNames.Clear

For X2 = 0 To lstApps.ListCount
        lstHwnd.AddItem lstApps.List(X2)
        lstHwndNames.AddItem lstNames.List(X2)
Next
End If
End Function


Private Sub Command1_Click(Index As Integer)
MakeNormal Command1(Index).Tag
End Sub

Private Sub Command1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If Button = 1 Then
    SetFGWindow Command1(Index).Tag, True
ElseIf Button = 2 Then
    SetFGWindow Command1(Index).Tag, False
End If
End Sub

Private Sub Form_Load()
Set Form = Form1
Set Bar = frmBar
Image1.ZOrder 0
WindowPos Me, 1
FadeIn Me, 255, 255
Command1(0).Top = Command1(0).Top - Command1(0).Height
fEnumWindows Me.lstApps
DoEvents
InitButtons
DoEvents
End Sub

Private Sub MacButton1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ReleaseCapture
SendMessage Me.hwnd, &HA1, 2, 0&
SideBarHeight
End Sub

Private Sub MacButton1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
SideBarHeight
End Sub

Private Sub Picture5_Click()
If Animation = True Then
FadeOut Me, 255, 0
End If
Unload Me
End Sub

Private Sub Picture6_Click()
Me.Tag = Bar.Addbutton(Me)
ShowWindow Me.hwnd, 0
End Sub

Private Sub Timer1_Timer()
fEnumWindows Me.lstApps
DoEvents
CheckApps
End Sub

Function InitButtons()
On Error Resume Next
For X2 = 0 To lstApps.ListCount - 1
        Load Command1(Command1.UBound + 1)
        Load Picture4(Picture4.UBound + 1)
        With Command1(Command1.UBound)
                .Top = Command1(Command1.UBound - 1).Top + Command1(Command1.UBound - 1).Height + 2
                .Caption = CheckCaption(lstNames.List(X2))
                .ToolTipText = lstNames.List(X2)
                Form.AddLine "+app+ " & .ToolTipText
                .Tag = lstApps.List(X2)
                .Visible = True
                .ZOrder 0
        End With
        With Picture4(Picture4.UBound)
                .Top = Command1(Command1.UBound).Top + 5
                .AutoRedraw = True
                .Visible = True
                .ZOrder 0
                Call DrawIcon(Picture4(Picture4.UBound).hdc, lstApps.List(X2), 0, 0)
        End With
DoEvents
Next
Timer1.Enabled = True
Form.AddLine "-Running Applications: " & lstApps.ListCount
SideBarHeight
End Function

Function Add_Remove_Button(Add As Boolean, Remove As Boolean, Optional Tag As Long, Optional Caption As String, Optional ToolTipText As String)
Dim ButtonID
If Add <> Remove Then
        If Add = True Then 'Add NEW Button
                Load Command1(Command1.UBound + 1)
                Load Picture4(Picture4.UBound + 1)
                With Command1(Command1.UBound)
                        .Top = Command1(Command1.UBound - 1).Top + Command1(Command1.UBound - 1).Height + 2
                        .Caption = Caption
                        .ToolTipText = ToolTipText
                        Form.AddLine "+ New App + " & .ToolTipText
                        .Tag = Tag
                        .Visible = True
                        .ZOrder 0
                End With
                With Picture4(Picture4.UBound)
                        .Top = Command1(Command1.UBound).Top + 5
                        .AutoRedraw = True
                        .Visible = True
                        .ZOrder 0
                        Call DrawIcon(Picture4(Picture4.UBound).hdc, Tag, 0, 0)
                End With
        Else ' REMOVE OLD BUTTON
                For X2 = 1 To Command1.UBound
                        If Command1(X2).Tag = Tag Then
                                ButtonID = X2
                                Exit For
                        End If
                Next
                For X2 = ButtonID To Command1.UBound - 1
                With Command1(X2)
                        .Caption = Command1(X2 + 1).Caption
                        .Tag = Command1(X2 + 1).Tag
                        .ToolTipText = Command1(X2 + 1).ToolTipText
                        Form.AddLine "+ Lost App + " & .ToolTipText
                        .Visible = True
                End With
                With Picture4(X2)
                        .Cls
                        .AutoRedraw = True
                        .Visible = True
                        Call DrawIcon(Picture4(X2).hdc, Command1(X2 + 1).Tag, 0, 0)
                End With
                Next
                Unload Command1(Command1.UBound)
                Unload Picture4(Picture4.UBound)
        End If
SideBarHeight
End If
End Function

Public Sub DrawIcon(hdc As Long, hwnd As Long, x As Integer, y As Integer)
ico = GetIcon(hwnd)
DrawIconEx hdc, x, y, ico, 16, 16, 0, 0, DI_NORMAL
End Sub

Public Function GetIcon(hwnd As Long) As Long
Call SendMessageTimeout(hwnd, WM_GETICON, 0, 0, 0, 1000, GetIcon)
If Not CBool(GetIcon) Then GetIcon = GetClassLong(hwnd, GCL_HICONSM)
If Not CBool(GetIcon) Then Call SendMessageTimeout(hwnd, WM_GETICON, 1, 0, 0, 1000, GetIcon)
If Not CBool(GetIcon) Then GetIcon = GetClassLong(hwnd, GCL_HICON)
If Not CBool(GetIcon) Then Call SendMessageTimeout(hwnd, WM_QUERYDRAGICON, 0, 0, 0, 1000, GetIcon)
End Function

Function CheckCaption(Text) As String
Dim TextLen
TextLen = 30
If Len(Text) > TextLen Then
Text = Left(Text, TextLen) & "..."
End If
CheckCaption = Text
End Function

Function SideBarHeight()
MacButton1.Height = Command1.UBound * (Command1(0).Height + 2) + 6
MacButton1.Height = Command1.UBound * (Command1(0).Height + 2) + 7
Me.Height = MacButton1.Height * 15
TransparentForm Me
End Function
