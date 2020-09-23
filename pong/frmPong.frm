VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0000C000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4230
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrKey 
      Interval        =   1
      Left            =   4320
      Top             =   3360
   End
   Begin VB.Timer tmrComputer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3840
      Top             =   3840
   End
   Begin VB.Timer tmrPong 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4320
      Top             =   3840
   End
   Begin VB.Shape shpBall 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label lblScorePlayer2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   165
   End
   Begin VB.Label lblScorePlayer1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   165
   End
   Begin VB.Label lblInstruct2 
      BackStyle       =   0  'Transparent
      Caption         =   "Press SPACE to serve         ESC to exit"
      Height          =   615
      Left            =   1560
      TabIndex        =   3
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label lblInstruct1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmPong.frx":0000
      Height          =   1035
      Left            =   1118
      TabIndex        =   2
      Top             =   720
      Width           =   2445
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblPlayer2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2280
      TabIndex        =   1
      Top             =   3960
      Width           =   120
   End
   Begin VB.Label lblPlayer1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2400
      TabIndex        =   0
      Top             =   0
      Width           =   120
   End
   Begin VB.Line lneNet 
      X1              =   0
      X2              =   4680
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Shape shpPad2 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   1920
      Top             =   3960
      Width           =   735
   End
   Begin VB.Shape shpPad1 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   2040
      Top             =   0
      Width           =   735
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Visible         =   0   'False
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intLeft As Integer
Dim intTop As Integer
Dim intDirection As Integer
Dim blnGame As Boolean
Dim intPlayerServe As Integer 'For player that will serve 0-lower player 1-upperplayer
Dim intMove As Integer
Dim blnHit As Boolean
Dim blnKey(255) As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    blnKey(KeyCode) = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    blnKey(KeyCode) = False
End Sub

Private Sub Form_Load()
'Position the Pad, the Ball and the Net
intPlayerServe = 0
PositionPlayer
PositionBall (intPlayerServe)
lneNet.X1 = 0
lneNet.Y1 = Me.Height / 2
lneNet.X2 = Me.Width
lneNet.Y2 = Me.Height / 2
blnGame = False
intMove = 50
blnHit = False

If intMode = 0 Then
    tmrComputer.Enabled = True
    lblPlayer1.Caption = ""
End If

'Set the default direction
intDirection = 1
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    PopupMenu mnuHelp
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Private Sub mnuAbout_Click()
MsgBox "Pongporongpongpong YALP v0.01" & vbNewLine & _
    "(Yet Another Lame Pong)" & vbNewLine & _
    "-Boring game" & vbNewLine & _
    "-Bad programming style" & vbNewLine & _
    "-Full of Bugs" & vbNewLine & _
    "-Hahahah" & vbNewLine & _
    "Created by Sec (www.secs81.tk)" & vbNewLine & vbNewLine & _
    "Special Thanx to:" & vbNewLine & _
    "Merri (Merri.net)" & vbNewLine & _
    "Cathoy", , "Pongporongpongpong"
End Sub

Private Sub tmrKey_Timer()
    If blnKey(65) Then 'A Key for Upper Player Left Direction
        If intMode = 1 Then
            If shpPad1.Left > 0 Then
                If intPlayerServe = 1 Then
                    If Not blnGame Then
                        shpBall.Left = shpBall.Left - 200
                    End If
                End If
                shpPad1.Left = shpPad1.Left - 200
                lblPlayer1.Left = lblPlayer1.Left - 200
            End If
        End If
    End If
    
    If blnKey(68) Then 'D Key for Upper Player Right Direction
        If intMode = 1 Then
            If (shpPad1.Left + shpPad1.Width) < Me.Width Then
                If intPlayerServe = 1 Then
                    If Not blnGame Then
                        shpBall.Left = shpBall.Left + 200
                    End If
                End If
                shpPad1.Left = shpPad1.Left + 200
                lblPlayer1.Left = lblPlayer1.Left + 200
            End If
        End If
    End If
        
    If blnKey(37) Then  'Arrow Key Left for Lower Player
        If shpPad2.Left > 0 Then
            If intPlayerServe = 0 Then
                If Not blnGame Then
                    shpBall.Left = shpBall.Left - 200
                End If
            End If
            shpPad2.Left = shpPad2.Left - 200
            lblPlayer2.Left = lblPlayer2.Left - 200
        End If
    End If
    
    If blnKey(39) Then 'Arrow Key Right for Lower Player
        If (shpPad2.Left + shpPad2.Width) < Me.Width Then
            If intPlayerServe = 0 Then
                If Not blnGame Then
                    shpBall.Left = shpBall.Left + 200
                End If
            End If
            shpPad2.Left = shpPad2.Left + 200
            lblPlayer2.Left = lblPlayer2.Left + 200
        End If
    End If
  
    If blnKey(32) Then
        tmrPong.Enabled = True
        blnGame = True
        lblInstruct1.Visible = False
        lblInstruct2.Visible = False
    End If
    
    If blnKey(27) Then
        MsgBox "www.secs81.tk", , "Pongporongpong"
        End
    End If
            
            
End Sub

Private Sub tmrComputer_Timer()
If (intDirection = 1 Or intDirection = 4) And shpBall.Top < Me.Height / 2 Then
    If shpPad1.Left + (shpPad1.Width / 2) > shpBall.Left + (shpBall.Width / 2) Then
        If (shpPad1.Left + (shpPad1.Width / 2)) - _
            (shpBall.Left + (shpBall.Width / 2)) < 200 Then
            shpPad1.Left = shpPad1.Left - ((shpPad1.Left + (shpPad1.Width / 2)) - (shpBall.Left + (shpBall.Width / 2)))
        Else
            shpPad1.Left = shpPad1.Left - 200
        End If
    ElseIf shpBall.Left + (shpBall.Width / 2) > shpPad1.Left + (shpPad1.Width / 2) Then
        If (shpBall.Left + (shpBall.Width / 2)) - _
            (shpPad1.Left + (shpPad1.Width / 2)) < 200 Then
            
            shpPad1.Left = shpPad1.Left + ((shpBall.Left + (shpBall.Width / 2)) - (shpPad1.Left + (shpPad1.Width / 2)))
        Else
            shpPad1.Left = shpPad1.Left + 200
        End If
    End If
End If

End Sub


Private Sub tmrPong_Timer()
Static intLoop As Integer

intLeft = shpBall.Left
intTop = shpBall.Top

'Move the ball
Select Case intDirection
    Case 1
        If shpBall.Top - (shpPad1.Top + shpPad1.Height) < intMove Then 'For the collision in the upper pad of the ball
            If blnHit Then
                shpBall.Left = shpBall.Left + intMove
                shpBall.Top = shpPad1.Top + shpPad1.Height
            Else
                shpBall.Left = shpBall.Left + intMove
                shpBall.Top = shpBall.Top - intMove
            End If
        ElseIf Me.Width - (shpBall.Left + shpBall.Width) < intMove Then 'For the collision in the right wall of the ball
            shpBall.Left = Me.Width - shpBall.Width
            shpBall.Top = shpBall.Top - intMove
        Else
            shpBall.Left = shpBall.Left + intMove
            shpBall.Top = shpBall.Top - intMove
        End If
    Case 2
        If shpPad2.Top - (shpBall.Top + shpBall.Height) < intMove Then 'For the collision of the lower pad
            If blnHit Then
                shpBall.Left = shpBall.Left + intMove
                shpBall.Top = shpPad2.Top - shpBall.Height
            Else
                shpBall.Left = shpBall.Left + intMove
                shpBall.Top = shpBall.Top + intMove
            End If
        ElseIf Me.Width - (shpBall.Left + shpBall.Width) < intMove Then 'For the collision to the right wall
            shpBall.Left = Me.Width - shpBall.Width
            shpBall.Top = shpBall.Top + intMove
        Else
            shpBall.Left = shpBall.Left + intMove
            shpBall.Top = shpBall.Top + intMove
        End If
    Case 3
        If shpPad2.Top - (shpBall.Top + shpBall.Height) < intMove Then 'For the collision of the lower pad
            If blnHit Then
                shpBall.Left = shpBall.Left - intMove
                shpBall.Top = shpPad2.Top - shpBall.Height
            Else
                shpBall.Left = shpBall.Left - intMove
                shpBall.Top = shpBall.Top + intMove
            End If
        ElseIf shpBall.Left < intMove Then  'For the collision to the left wall
            shpBall.Left = 0
            shpBall.Top = shpBall.Top + intMove
        Else
            shpBall.Left = shpBall.Left - intMove
            shpBall.Top = shpBall.Top + intMove
        End If
    Case 4
        If shpBall.Top - (shpPad1.Top + shpPad1.Height) < intMove Then 'For the collision of the upper pad
            If blnHit Then
                shpBall.Left = shpBall.Left - intMove
                shpBall.Top = shpPad1.Top + shpPad1.Height
            Else
                shpBall.Left = shpBall.Left - intMove
                shpBall.Top = shpBall.Top - intMove
            End If
        ElseIf shpBall.Left < intMove Then  'For the collision of ther left wall
            shpBall.Left = 0
            shpBall.Top = shpBall.Top - intMove
        Else
            shpBall.Left = shpBall.Left - intMove
            shpBall.Top = shpBall.Top - intMove
        End If
End Select

'COLLISSION DETECTION: WALL, PAD AND BASE
If shpBall.Left <= 0 Then  'Hit the left wall

    If shpBall.Top <= (shpPad1.Top + shpPad1.Height) Then
        intDirection = 2
    ElseIf ((shpBall.Top + shpBall.Height) >= shpPad2.Top) Then 'When the ball hit the third corner
        intDirection = 1
    Else
        If (intLeft > shpBall.Left) And (intTop > shpBall.Top) Then
            intDirection = 1
        ElseIf (intLeft > shpBall.Left) And (intTop < shpBall.Top) Then
            intDirection = 2
        End If
    End If
    Beep
ElseIf (shpBall.Left + shpBall.Width) >= Me.Width Then 'Hit the right wall
    
    If shpBall.Top <= (shpPad1.Top + shpPad1.Height) Then 'When the ball hit the second corner
        intDirection = 3
    ElseIf (shpBall.Top + shpBall.Height) >= shpPad2.Top Then 'When the ball hit the fourth corner
        intDirection = 4
    Else
        If (intLeft < shpBall.Left) And (intTop < shpBall.Top) Then
            intDirection = 3
        ElseIf (intLeft < shpBall.Left) And (intTop > shpBall.Top) Then
            intDirection = 4
        End If
        Beep
    End If
    
ElseIf shpBall.Top <= (shpPad1.Top + shpPad1.Height) Then 'Hit the upper pad
    If (shpBall.Left >= shpPad1.Left And shpBall.Left <= (shpPad1.Left + shpPad1.Width)) _
        Or ((shpBall.Left + shpBall.Width) >= shpPad1.Left And (shpBall.Left + shpBall.Width) <= _
        (shpPad1.Left + shpPad1.Width)) Then 'Ball in the horizontal hit area of the pad
        If intTop > shpBall.Top And intLeft < shpBall.Left Then
            blnHit = True
            intDirection = 2
        ElseIf intTop > shpBall.Top And intLeft > shpBall.Left Then
            blnHit = True
            intDirection = 3
        End If
        intLoop = intLoop + 1
        Beep
    Else
        blnHit = False
    End If
    
    If shpBall.Top < 0 Then 'Hit the upper base
        tmrPong.Enabled = False
        intLoop = 0
        intMove = 50
        intPlayerServe = 1
        PositionPlayer
        PositionBall intPlayerServe
        blnGame = False
        intDirection = 3
        lblScorePlayer2.Caption = Val(Trim(lblScorePlayer2.Caption)) + 1
        Beep
    End If
ElseIf (shpBall.Top + shpBall.Height) >= shpPad2.Top Then 'Hit the lower pad
    If (shpBall.Left >= shpPad2.Left And shpBall.Left <= (shpPad2.Left + shpPad2.Width)) _
        Or ((shpBall.Left + shpBall.Width) >= shpPad2.Left And (shpBall.Left + shpBall.Width) <= _
        (shpPad2.Left + shpPad2.Width)) Then 'Ball in the horizontal hit area of the pad
        If intTop < shpBall.Top And intLeft > shpBall.Left Then
            blnHit = True
            intDirection = 4
        ElseIf intTop < shpBall.Top And intLeft < shpBall.Left Then
            blnHit = True
            intDirection = 1
        End If
        intLoop = intLoop + 1
        Beep
    Else
        blnHit = False
    End If
    If (shpBall.Top + shpBall.Height) > Me.Height Then 'Hit the lower base
        tmrPong.Enabled = False
        intMove = 50
        intLoop = 0
        intPlayerServe = 0
        PositionPlayer
        PositionBall intPlayerServe
        blnGame = False
        intDirection = 1
        lblScorePlayer1.Caption = Val(Trim(lblScorePlayer1.Caption)) + 1
        Beep
    End If
End If 'End of Collision Detectioin
If intLoop = 10 Then
    intMove = intMove + 50
    intLoop = 0
End If
DoEvents
End Sub

Private Sub PositionBall(intPlayer As Integer)
    If intPlayer = 0 Then
        shpBall.Left = (Me.Width - shpBall.Width) / 2
        shpBall.Top = Me.Height - (shpBall.Height + shpPad2.Height)
    Else
        shpBall.Left = (Me.Width - shpBall.Width) / 2
        shpBall.Top = shpPad1.Top + shpPad1.Height
    End If
End Sub

Private Sub PositionPlayer()
shpPad1.Top = 0
shpPad1.Left = (Me.Width - shpPad1.Width) / 2
lblPlayer1.Left = (shpPad1.Left + (shpPad1.Width / 2)) - (lblPlayer1.Width / 2)
lblPlayer1.Top = (shpPad1.Top + (shpPad1.Height / 2)) - (lblPlayer1.Height / 2)
shpPad2.Top = Me.Height - shpPad2.Height
shpPad2.Left = (Me.Width - shpPad2.Width) / 2
lblPlayer2.Left = (shpPad2.Left + (shpPad2.Width / 2)) - (lblPlayer2.Width / 2)
lblPlayer2.Top = (shpPad2.Top + (shpPad2.Height / 2)) - (lblPlayer2.Height / 2)
End Sub
