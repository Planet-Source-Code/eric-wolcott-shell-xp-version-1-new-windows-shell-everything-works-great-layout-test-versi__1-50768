VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuDesktop 
      Caption         =   "mnuDesktop"
      Begin VB.Menu mnuIcon 
         Caption         =   "mnuIcon"
         Begin VB.Menu mnuOpen 
            Caption         =   "Open"
         End
         Begin VB.Menu mnuRename 
            Caption         =   "Rename"
         End
         Begin VB.Menu mnuDelete 
            Caption         =   "Delete"
         End
         Begin VB.Menu mnuProperties 
            Caption         =   "Properties"
         End
      End
      Begin VB.Menu mnuBackround 
         Caption         =   "mnuBackround"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cIcon As Long
