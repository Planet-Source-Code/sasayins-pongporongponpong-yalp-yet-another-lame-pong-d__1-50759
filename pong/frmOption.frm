VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Pongporongpongpong YALP v0.01"
   ClientHeight    =   1365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3060
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   1365
   ScaleWidth      =   3060
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1043
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.OptionButton optMultiPlayer 
      Caption         =   "Two Player (Human vs Human)"
      Height          =   195
      Left            =   83
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.OptionButton optSinglePlayer 
      Caption         =   "Single Player (Human vs Computer)"
      Height          =   195
      Left            =   83
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
If optSinglePlayer.Value = True Then
    intMode = 0
Else
    intMode = 1
End If
Me.Hide
Form1.Show
End Sub
