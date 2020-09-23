VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   3720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4200
   LinkTopic       =   "Form3"
   ScaleHeight     =   3720
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.API API 
      Height          =   480
      Left            =   1560
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   3840
      Picture         =   "Form3.frx":0000
      Top             =   105
      Width           =   240
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   480
      Top             =   1320
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Agent As IAgentCtlCharacterEx  ' Make the name "Genie" An Agent Character
Public EditListIndex As Integer
Function LoadAgent()

Dim xload As String
Agent1.Characters.Load "James", "James.acs"
Set Agent = Agent1.Characters("James")
Agent.LanguageID = &H409
Agent.Left = ((Screen.Width - Agent.Width) / 2) / 18
Agent.Top = ((Screen.Height - Agent.Height) / 2) / 15
Me.Left = Agent.Left * 15
Me.Top = Agent.Top * 15
Me.Width = Agent.Width * 15
Me.Height = Agent.Height * 15
Agent.Show
Agent.Play "Greet"
Agent.Speak "Hello, And Welcome To The, Shell XP Program For Windows, Written By Zach Szafran." 'Say something
'Agent.Left = 200
'Agent.Top = 200
xload = GetSetting(App.EXEName, "LoadIntro", "Load", True)
Loadintro xload
End Function

Function Loadintro(Load As String)
If Load = "True" Then
Agent.Speak "If you look"
Agent.Play "gestureright"
Agent.Speak "To your right, you can see the default shortcuts already created"
Agent.Speak "Right click on an icon to display it's menu"
Agent.Speak "But please wait ontill i'm finished"
Agent.Speak "Left click and hold the icon to drag it to another position"
Agent.Play "GestureUp"
Agent.Speak "Slightly above me is the location of other icons"
Agent.Speak "These icons are ones of importance"
Agent.Speak "Click the menu icon to display this program's menu"
Agent.Speak "In this menu the following are located,,, bandwidth Usage, CPU Usage, task-bar, and Exit Program"
Agent.Play "gestureright"
Agent.Speak "Also to you upper right the Task-Bar is located"
Agent.Speak "This task-bar can be use to view you runing applications"
Agent.Speak "Simply click on the application you wish to restore"
Agent.Play "GestureDown"
Agent.Speak "The blue bar to the bottom contains you clock, and your Shell-XP running tasks"
Agent.Play "Greet"
Agent.Speak "I will be Popping up through out this program to give you tips"
Agent.Speak "Good luck"
Agent.Play "hide"
Me.Visible = False
End If
End Function


Private Sub Form_Load()
If API.Path_Exist("C:\WINNT\msagent\chars") = 1 Then
API.Copy_File App.path & "\agents\james.acs", "C:\WINNT\msagent\chars"
LoadAgent
ElseIf API.Path_Exist("C:\WINDOWS\msagent\chars") = 1 Then
API.Copy_File App.path & "\agents\james.acs", "C:\WINDOWS\msagent\chars"
LoadAgent
Else
Unload Me
End If
End Sub

Private Sub Image7_Click()

End Sub
