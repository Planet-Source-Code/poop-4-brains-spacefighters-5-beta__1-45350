VERSION 5.00
Begin VB.Form frmNew 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spacefighters - New Game"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
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
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   370
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNew 
      Caption         =   "New Game"
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   3240
      Width           =   1695
   End
   Begin VB.PictureBox picPlayers 
      Height          =   615
      Left            =   120
      Picture         =   "frmNew.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   5235
      TabIndex        =   2
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton cmdP 
         Caption         =   "4"
         Height          =   375
         Index           =   2
         Left            =   3720
         TabIndex        =   15
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdP 
         Caption         =   "3"
         Height          =   375
         Index           =   1
         Left            =   2520
         TabIndex        =   14
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdP 
         Caption         =   "2"
         Height          =   375
         Index           =   0
         Left            =   1320
         TabIndex        =   13
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Players:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.PictureBox picShip 
      Height          =   2295
      Left            =   120
      Picture         =   "frmNew.frx":3A9C2
      ScaleHeight     =   2235
      ScaleWidth      =   5235
      TabIndex        =   1
      Top             =   840
      Width           =   5295
      Begin VB.ComboBox cmbShip 
         Height          =   390
         Index           =   3
         ItemData        =   "frmNew.frx":75384
         Left            =   3000
         List            =   "frmNew.frx":75391
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1680
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ComboBox cmbShip 
         Height          =   390
         Index           =   2
         ItemData        =   "frmNew.frx":753AA
         Left            =   3000
         List            =   "frmNew.frx":753B7
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1200
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ComboBox cmbShip 
         Height          =   390
         Index           =   1
         ItemData        =   "frmNew.frx":753D0
         Left            =   3000
         List            =   "frmNew.frx":753DD
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   720
         Width           =   2055
      End
      Begin VB.ComboBox cmbShip 
         Height          =   390
         Index           =   0
         ItemData        =   "frmNew.frx":753F6
         Left            =   3000
         List            =   "frmNew.frx":75403
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   2055
      End
      Begin VB.Shape shpShip 
         Height          =   375
         Index           =   3
         Left            =   120
         Top             =   1680
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.Shape shpShip 
         Height          =   375
         Index           =   2
         Left            =   120
         Top             =   1200
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.Shape shpShip 
         Height          =   375
         Index           =   1
         Left            =   120
         Top             =   720
         Width           =   4935
      End
      Begin VB.Shape shpShip 
         Height          =   375
         Index           =   0
         Left            =   120
         Top             =   240
         Width           =   4935
      End
      Begin VB.Label lblShip 
         BackStyle       =   0  'Transparent
         Caption         =   "Player 4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label lblShip 
         BackStyle       =   0  'Transparent
         Caption         =   "Player 3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label lblShip 
         BackStyle       =   0  'Transparent
         Caption         =   "Player 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label lblShip 
         BackStyle       =   0  'Transparent
         Caption         =   "Player 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   3240
      Width           =   1575
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Players As Integer

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdNew_Click()
Dim i As Long

For i = 1 To 4
    P(i).Act = False
Next i

For i = 1 To 4
    If lblShip(i - 1).Visible = True Then

        P(i).Act = True
        
        P(i).Xs = 0
        P(i).Ys = 0
        
        P(i).Tag = i
        P(i).Ty = cmbShip(i - 1).ListIndex

        P(i).MHP = 50
        P(i).HP = 50
        P(i).Ammo = 10

        Select Case i
        Case 1
            P(i).x = 10
            P(i).y = 10
        Case 2
            P(i).x = frmGame.Board.ScaleWidth - 60
            P(i).y = 10
        Case 3
            P(i).x = 10
            P(i).y = frmGame.Board.ScaleHeight - 60
        Case 4
            P(i).x = frmGame.Board.ScaleWidth - 60
            P(i).y = frmGame.Board.ScaleHeight - 60
        End Select
    End If
Next i

For i = 1 To 10
    A(i).Act = False
Next i

For i = 1 To 15
    S(i).Act = False
Next i

frmGame.cmdNew.Enabled = False
frmGame.cmdOptions.Enabled = False
frmGame.cmdExit.Enabled = False

Unload Me
Running = True
frmGame.GameLoop
End Sub

Private Sub cmdP_Click(Index As Integer)
Dim i As Long

For i = 0 To 3
    cmbShip(i).Visible = False
    shpShip(i).Visible = False
    lblShip(i).Visible = False
Next i

For i = 0 To Index + 1
    cmbShip(i).Visible = True
    shpShip(i).Visible = True
    lblShip(i).Visible = True
Next i

End Sub

Private Sub Form_Load()
Dim i As Long

cmdP_Click (0)

For i = 0 To 3
    cmbShip(i).ListIndex = 0
Next i
End Sub

