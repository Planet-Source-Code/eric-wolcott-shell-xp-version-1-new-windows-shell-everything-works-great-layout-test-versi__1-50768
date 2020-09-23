VERSION 5.00
Begin VB.Form frmConsole 
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5580
   Icon            =   "frmConsole.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   3090
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Left            =   645
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   840
      Width           =   4185
   End
End
Attribute VB_Name = "frmConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private FormMe As NewForm2
'Function LoadForm()
'UserControl41.Width = Me.Width
'UserControl41.Top = Me.Height - UserControl41.Height
'UserControl41.Left = 0
'
'UserControl31.Top = UserControl11.Height
'UserControl31.Left = Me.Width - UserControl31.Width
'UserControl31.Height = Me.Height - UserControl41.Height - UserControl11.Height
'
'UserControl21.Top = UserControl11.Height
'UserControl21.Left = 0
'UserControl21.Height = Me.Height - UserControl41.Height - UserControl11.Height
'
'UserControl11.Top = 0
'UserControl11.Left = 0
'UserControl11.Width = Me.Width
'
'FormMe.Top = UserControl11.Height
'FormMe.Left = UserControl21.Width
'FormMe.Height = Me.Height - UserControl11.Height - UserControl41.Height
'FormMe.Width = Me.Width - UserControl21.Width - UserControl31.Width
'End Function

Private Sub Form_Load()
'Set FormMe = NewForm
'LoadForm
'txtConsole.Top = FormMe.Top
'txtConsole.Left = FormMe.Left
'txtConsole.Height = FormMe.Height
'txtConsole.Width = FormMe.Width
End Sub
