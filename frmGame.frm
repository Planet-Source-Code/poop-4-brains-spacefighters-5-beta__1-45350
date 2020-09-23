VERSION 5.00
Begin VB.Form frmGame 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spacefighters 5"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   514
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   681
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Panel 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   3
      Left            =   7680
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   8
      Top             =   1440
      Width           =   2415
   End
   Begin VB.PictureBox Panel 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   2
      Left            =   5160
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   7
      Top             =   1440
      Width           =   2415
   End
   Begin VB.PictureBox Panel 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   1
      Left            =   2640
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   6
      Top             =   1440
      Width           =   2415
   End
   Begin VB.PictureBox Panel 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   120
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   5
      Top             =   1440
      Width           =   2415
   End
   Begin VB.PictureBox Board 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   0
      ScaleHeight     =   393
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   681
      TabIndex        =   1
      Top             =   1800
      Width           =   10215
   End
   Begin VB.PictureBox Menu 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   0
      Picture         =   "frmGame.frx":0000
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   681
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit Game"
         Height          =   375
         Left            =   7800
         TabIndex        =   4
         Top             =   840
         Width           =   2295
      End
      Begin VB.CommandButton cmdOptions 
         Caption         =   "Options"
         Height          =   375
         Left            =   7800
         TabIndex        =   3
         Top             =   480
         Width           =   2295
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New Game"
         Height          =   375
         Left            =   7800
         TabIndex        =   2
         Top             =   120
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Board_Click()
If Running = True Then
    If Paused = True Then
        Paused = False
        frmGame.cmdNew.Enabled = False
        frmGame.cmdOptions.Enabled = False
        frmGame.cmdExit.Enabled = False
    Else
        Paused = True
        frmGame.cmdNew.Enabled = True
        frmGame.cmdOptions.Enabled = True
        frmGame.cmdExit.Enabled = True
    End If
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdNew_Click()
Load frmNew
frmNew.Show
End Sub

Function GameLoop()
Dim C As Long, y As Long

Do Until Running = False

        If C > Speed Then
            C = 0

            Board.Cls

            If Paused = True Then

                Board.ForeColor = vbRed
                Board.FontSize = 32
                Board.CurrentX = Board.ScaleWidth \ 2 - Board.TextWidth("PAUSED") \ 2
                Board.CurrentY = Board.ScaleHeight \ 2 - Board.TextHeight("|") \ 2
                Board.Print "Paused"
            
            Else

                DoAllKeys
                MoveObjects

                For y = 1 To 4
                    If P(y).Act = True Then
                        BitBlt Board.hdc, P(y).x, P(y).y, 50, 45, frmGfx.ShipM(P(y).Ty).hdc, P(y).Rot * 50, 0, SRCAND
                        BitBlt Board.hdc, P(y).x, P(y).y, 50, 45, frmGfx.ShipS(P(y).Ty).hdc, P(y).Rot * 50, 0, SRCINVERT
                    End If
                Next y


                For y = 1 To 10
                    If A(y).Act = True Then
                        BitBlt Board.hdc, A(y).x, A(y).y, 50, 45, frmGfx.AstM(P(y).Ty).hdc, A(y).Rot * 50, 0, SRCAND
                        BitBlt Board.hdc, A(y).x, A(y).y, 50, 45, frmGfx.AstS(P(y).Ty).hdc, A(y).Rot * 50, 0, SRCINVERT
                    End If
                Next y

                For y = 1 To 15
                    If S(y).Act = True Then
                        BitBlt Board.hdc, S(y).x, S(y).y, 15, 15, frmGfx.ShM.hdc, 0, 0, SRCAND
                        BitBlt Board.hdc, S(y).x, S(y).y, 15, 15, frmGfx.ShS.hdc, 0, 0, SRCINVERT
                    End If
                Next y
                
                For y = 1 To 4
                    Panel(y - 1).Cls

                    If P(y).Act = True Then
                        Panel(y - 1).Cls
                        Panel(y - 1).CurrentX = 10
                        Panel(y - 1).CurrentY = 2
                        Panel(y - 1).Print "Player " & y & ": "
                        Panel(y - 1).Line (Panel(y - 1).TextWidth("Player " & y & ": ") + 11, 5)-(Panel(y - 1).TextWidth("Player " & y & ": ") + 11 + 50, 20), vbRed, BF
                        Panel(y - 1).Line (Panel(y - 1).TextWidth("Player " & y & ": ") + 11, 5)-(Panel(y - 1).TextWidth("Player " & y & ": ") + 11 + P(y).HP, 20), vbGreen, BF
                    End If
                Next y

            End If

        Else
            C = C + 1

        End If

    DoEvents

Loop

End Function

Private Sub cmdOptions_Click()
Load frmOptions
frmOptions.Show
End Sub

Private Sub Form_Load()
Dim T As Long

Speed = GetSetting("SpaceFighters5", "Speed", "Control", 5000)

GetAsyncKeyState (vbKeySpace)

Open App.Path & "\keys.dat" For Input As #1

For T = 1 To 4
    Input #1, Pc(T).Forward, Pc(T).Left, Pc(T).Right, Pc(T).Shoot
Next T

Close #1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Running = False
Paused = False

SaveSetting "SpaceFighters5", "Speed", "Control", Speed

Unload frmNew
Unload frmGfx
Unload frmOptions
End Sub
