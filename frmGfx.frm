VERSION 5.00
Begin VB.Form frmGfx 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   ScaleHeight     =   350
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   563
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox ShipS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   2
      Left            =   120
      Picture         =   "frmGfx.frx":0000
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VB.PictureBox ShipM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   2
      Left            =   120
      Picture         =   "frmGfx.frx":D332
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   10
      Top             =   1680
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VB.PictureBox ShS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   7440
      Picture         =   "frmGfx.frx":1A664
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox ShM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   7440
      Picture         =   "frmGfx.frx":1A976
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox AstM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   1
      Left            =   120
      Picture         =   "frmGfx.frx":1AC88
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   7
      Top             =   4440
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.PictureBox AstS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   1
      Left            =   120
      Picture         =   "frmGfx.frx":2B476
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   6
      Top             =   3600
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.PictureBox AstM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   0
      Left            =   120
      Picture         =   "frmGfx.frx":3BC64
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   5
      Top             =   4320
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.PictureBox AstS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   0
      Left            =   120
      Picture         =   "frmGfx.frx":4C452
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   4
      Top             =   3480
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.PictureBox ShipM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   1
      Left            =   120
      Picture         =   "frmGfx.frx":5CC40
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VB.PictureBox ShipM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   0
      Left            =   120
      Picture         =   "frmGfx.frx":69F72
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VB.PictureBox ShipS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   1
      Left            =   120
      Picture         =   "frmGfx.frx":772A4
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VB.PictureBox ShipS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   0
      Left            =   120
      Picture         =   "frmGfx.frx":845D6
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   6000
   End
End
Attribute VB_Name = "frmGfx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
